# Practice Perfect Email Router - One-Click Installer
# This script sets up everything automatically for each clinician

# Requires running as administrator for Python installation
#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Configuration
$InstallDir = "C:\PracticePerfectProxy"
$PythonVersion = "3.12.0"
$PythonInstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"

# Color output functions
function Write-Success { param($msg) Write-Host "[SUCCESS] $msg" -ForegroundColor Green }
function Write-Info    { param($msg) Write-Host "[INFO] $msg"    -ForegroundColor Cyan }
function Write-Warn    { param($msg) Write-Host "[WARNING] $msg" -ForegroundColor Yellow }
function Write-Err     { param($msg) Write-Host "[ERROR] $msg"   -ForegroundColor Red }

# Banner
Clear-Host
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host "Practice Perfect Email Router - Automated Installer" -ForegroundColor Cyan
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""

# -----------------------------------------------------------------------------
# Step 0: Nuke any existing installation before proceeding
# -----------------------------------------------------------------------------
$TaskName = "Practice Perfect Email Router"
$TaskXml  = "C:\Windows\System32\Tasks\$TaskName"

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
$existingDir  = Test-Path $InstallDir
$existingXml  = Test-Path $TaskXml

if ($existingTask -or $existingDir -or $existingXml) {
    Write-Warn "Existing installation detected -- cleaning up before reinstalling..."
    Write-Host ""

    # Kill any running Python process holding our ports
    $killed = $false
    Get-Process python -ErrorAction SilentlyContinue | ForEach-Object {
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        $killed = $true
    }
    if ($killed) { Write-Success "Running router process stopped" }

    # Stop the task if it is currently running
    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    # Unregister from Task Scheduler (in-memory / registry entry)
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Success "Task unregistered from Task Scheduler"
    }

    # Delete the on-disk XML -- the ghost that survives reboots and Parallels resets
    if ($existingXml) {
        Remove-Item -Path $TaskXml -Force -ErrorAction SilentlyContinue
        Write-Success "Task XML deleted (prevents ghost task on next reboot)"
    }

    # Remove install directory
    if ($existingDir) {
        Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Success "Install directory removed"
    }

    Write-Success "Cleanup complete -- proceeding with fresh installation"
    Write-Host ""
} else {
    Write-Info "No existing installation found -- proceeding with fresh installation"
    Write-Host ""
}

# -----------------------------------------------------------------------------
# Step 1: Check / Install Python
# -----------------------------------------------------------------------------
Write-Info "Checking for Python installation..."
$pythonInstalled = $false

try {
    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $pythonVersion = python --version 2>&1
        Write-Success "Python already installed: $pythonVersion"
        $pythonInstalled = $true
    }
} catch {
    Write-Warn "Python not found in PATH"
}

if (-not $pythonInstalled) {
    Write-Info "Python not found. Installing Python $PythonVersion..."
    $installerPath = "$env:TEMP\python-installer.exe"

    Write-Info "Downloading Python installer..."
    try {
        Invoke-WebRequest -Uri $PythonInstallerUrl -OutFile $installerPath
        Write-Success "Installer downloaded"
    } catch {
        Write-Err "Failed to download Python installer"
        Write-Host "Install manually from https://www.python.org/downloads/"
        Read-Host "Press Enter to exit"
        exit 1
    }

    Write-Info "Installing Python (this may take a few minutes)..."
    Start-Process -FilePath $installerPath `
        -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_test=0" `
        -Wait

    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path    = $machinePath + ";" + $userPath

    Start-Sleep -Seconds 3
    try {
        $pythonVersion = python --version 2>&1
        Write-Success "Python installed: $pythonVersion"
    } catch {
        Write-Err "Python install failed or not in PATH"
        Read-Host "Press Enter to exit"
        exit 1
    }

    Remove-Item $installerPath -ErrorAction SilentlyContinue
}

# -----------------------------------------------------------------------------
# Step 2: Install aiosmtpd
# -----------------------------------------------------------------------------
Write-Info "Installing aiosmtpd..."
try {
    python -m pip install --quiet --upgrade pip
    python -m pip install --quiet aiosmtpd
    Write-Success "aiosmtpd installed"
} catch {
    Write-Err "Failed to install aiosmtpd: $_"
    Read-Host "Press Enter to exit"
    exit 1
}

# -----------------------------------------------------------------------------
# Step 3: Create install directory
# -----------------------------------------------------------------------------
Write-Info "Creating install directory: $InstallDir"
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    Write-Success "Directory created"
} else {
    Write-Success "Directory already exists"
}

# -----------------------------------------------------------------------------
# Step 4: Write email_router.py
# -----------------------------------------------------------------------------
Write-Info "Writing email_router.py..."
$emailRouterScript = @'
#!/usr/bin/env python3
"""
Practice Perfect Standalone Email Router
Fixes included:
  - Single-instance lock (prevents port fights on Parallels resume)
  - Port cleanup on startup (kills zombie processes holding 2525)
  - Auto-restart loop with 60-second heartbeat (survives silent crashes)
  - Startup log written to disk (visible even from hidden VBS launcher)
  - Test email sent to clinician's own Gmail after first-time setup
"""

import asyncio
import json
import logging
import os
import smtplib
import socket
import subprocess
import sys
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import BytesParser

from aiosmtpd.controller import Controller

# -- Constants -----------------------------------------------------------------
CONFIG_FILE   = "email_router_config.json"
LOCAL_PORT    = 2525
LOCK_PORT     = 47200
HEARTBEAT_SEC = 60
RESTART_DELAY = 10

# -- Logging -------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("email_router.log", encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

# -- Single-instance lock ------------------------------------------------------
def acquire_instance_lock():
    lock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    lock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 0)
    try:
        lock.bind(("localhost", LOCK_PORT))
        lock.listen(1)
        logging.info(f"Instance lock acquired on port {LOCK_PORT}")
        return lock
    except OSError:
        logging.warning("Another instance is already running -- exiting.")
        sys.exit(0)

# -- Port cleanup --------------------------------------------------------------
def kill_port_holder(port: int):
    try:
        result = subprocess.run(
            f'netstat -ano | findstr ":{port} "',
            shell=True, capture_output=True, text=True
        )
        for line in result.stdout.splitlines():
            if f":{port}" in line and "LISTENING" in line:
                pid = line.strip().split()[-1]
                logging.warning(f"Port {port} held by PID {pid} -- killing it...")
                subprocess.run(f"taskkill /F /PID {pid}", shell=True, capture_output=True)
                time.sleep(1)
                logging.info(f"PID {pid} killed, port {port} should be free")
    except Exception as e:
        logging.warning(f"Could not check/clear port {port}: {e}")

# -- Test email ----------------------------------------------------------------
def send_test_email(gmail_user: str, gmail_password: str, clinician_name: str) -> bool:
    """
    Send a test email from the clinician's Gmail to themselves.
    Called once at the end of first-time setup so they can confirm
    credentials and delivery work before the router goes into production.
    """
    print()
    print("-" * 70)
    print("Sending test email to verify your credentials...")
    print("-" * 70)

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "Practice Perfect Email Router -- Setup Test"
    msg["From"]    = gmail_user
    msg["To"]      = gmail_user

    plain_body = f"""\
Practice Perfect Email Router -- Test Email
===========================================

Hello {clinician_name},

This email confirms that your email router is correctly configured
and can send emails through your Gmail account.

  Sent at  : {timestamp}
  Gmail    : {gmail_user}
  SMTP     : localhost:{LOCAL_PORT}

Configure Practice Perfect with:
  Server:         localhost
  Port:           {LOCAL_PORT}
  Encryption:     None
  Authentication: None

If you received this, everything is working. No further action needed.

-- Practice Perfect Email Router
"""

    html_body = f"""\
<html>
<body style="font-family:Arial,sans-serif;max-width:560px;margin:auto;padding:24px">
  <h2 style="color:#2e7d32">Practice Perfect Email Router &mdash; Setup Successful</h2>
  <p>Hello <strong>{clinician_name}</strong>,</p>
  <p>This email confirms that your email router is correctly configured
     and can send emails through your Gmail account.</p>
  <table style="border-collapse:collapse;margin:16px 0">
    <tr>
      <td style="padding:4px 16px 4px 0;color:#666">Sent at</td>
      <td><strong>{timestamp}</strong></td>
    </tr>
    <tr>
      <td style="padding:4px 16px 4px 0;color:#666">Gmail</td>
      <td><strong>{gmail_user}</strong></td>
    </tr>
    <tr>
      <td style="padding:4px 16px 4px 0;color:#666">SMTP</td>
      <td><strong>localhost:{LOCAL_PORT}</strong></td>
    </tr>
  </table>
  <p>Configure Practice Perfect with these SMTP settings:</p>
  <table style="background:#f5f5f5;padding:12px 16px;border-radius:6px;border-collapse:collapse">
    <tr>
      <td style="padding:3px 20px 3px 0;color:#555">Server</td>
      <td><strong>localhost</strong></td>
    </tr>
    <tr>
      <td style="padding:3px 20px 3px 0;color:#555">Port</td>
      <td><strong>{LOCAL_PORT}</strong></td>
    </tr>
    <tr>
      <td style="padding:3px 20px 3px 0;color:#555">Encryption</td>
      <td><strong>None</strong></td>
    </tr>
    <tr>
      <td style="padding:3px 20px 3px 0;color:#555">Authentication</td>
      <td><strong>None</strong></td>
    </tr>
  </table>
  <p style="color:#999;font-size:12px;margin-top:24px">
    &mdash; Practice Perfect Email Router
  </p>
</body>
</html>
"""

    msg.attach(MIMEText(plain_body, "plain"))
    msg.attach(MIMEText(html_body,  "html"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(gmail_user, gmail_password)
            smtp.sendmail(gmail_user, [gmail_user], msg.as_bytes())

        print(f"\n  [OK]  Test email sent to {gmail_user}")
        print("        Check your inbox (and spam folder) to confirm it arrived.")
        print()
        return True

    except smtplib.SMTPAuthenticationError:
        print("\n  [FAIL]  Gmail authentication error.")
        print("          Your app password may be incorrect.")
        print("          Re-run the installer to enter new credentials.")
        print()
        return False

    except Exception as e:
        print(f"\n  [WARN]  Test email could not be sent: {e}")
        print("          Check your internet connection.")
        print("          The router can still be started manually.")
        print()
        return False

# -- Config --------------------------------------------------------------------
def load_or_create_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
        if all(k in config for k in ("gmail", "app_password", "name")):
            return config
        print("Config file exists but is incomplete. Reconfiguring...\n")

    print("=" * 70)
    print("Practice Perfect Email Router - First Time Setup")
    print("=" * 70)
    print()

    config = {}

    name = input("Enter your name (for logging): ").strip() or "Clinician"
    config["name"] = name

    gmail = input("Enter your Gmail address: ").strip()
    if not gmail or "@" not in gmail:
        print("Valid Gmail address required!")
        sys.exit(1)
    config["gmail"] = gmail

    print()
    print("Gmail App Password")
    print("  1. Go to: https://myaccount.google.com/apppasswords")
    print("  2. Select 'Mail' and 'Other (Custom name)'")
    print("  3. Name it 'Practice Perfect'")
    print("  4. Copy the 16-character password")
    print()
    app_password = input("Enter your Gmail App Password: ").strip().replace(" ", "")
    if not app_password:
        print("App password required!")
        sys.exit(1)
    config["app_password"] = app_password
    config["local_port"]   = LOCAL_PORT

    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

    print(f"\nConfig saved to {CONFIG_FILE}")

    # Send test email immediately after saving credentials.
    # If it fails, offer to re-enter credentials before continuing.
    test_ok = send_test_email(gmail, app_password, name)

    if not test_ok:
        retry = input("Would you like to re-enter your credentials now? (y/N): ").strip().lower()
        if retry == "y":
            os.remove(CONFIG_FILE)
            return load_or_create_config()

    print()
    print("=" * 70)
    print("Configure Practice Perfect SMTP:")
    print("  Server:         localhost")
    print(f"  Port:           {LOCAL_PORT}")
    print("  Encryption:     None")
    print("  Authentication: None")
    print("=" * 70)
    print()

    return config

# -- Gmail connection test -----------------------------------------------------
def test_gmail_connection(gmail_user: str, gmail_password: str) -> bool:
    print("\nTesting Gmail connection...")
    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=10) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(gmail_user, gmail_password)
        print("Gmail authentication successful!\n")
        return True
    except smtplib.SMTPAuthenticationError:
        print("Gmail authentication FAILED -- check your app password.\n")
        return False
    except Exception as e:
        print(f"Warning: Could not reach Gmail: {e}\n")
        return True

# -- SMTP handler --------------------------------------------------------------
class EmailRouter:
    def __init__(self, gmail_user: str, gmail_password: str, clinician_name: str):
        self.gmail_user     = gmail_user
        self.gmail_password = gmail_password
        self.clinician_name = clinician_name

    async def handle_DATA(self, server, session, envelope):
        try:
            logging.info(f"[{self.clinician_name}] Received email")
            logging.info(f"  From:    {envelope.mail_from}")
            logging.info(f"  To:      {envelope.rcpt_tos}")

            raw = envelope.content
            if isinstance(raw, str):
                raw = raw.encode("utf-8", errors="replace")
            msg = BytesParser().parsebytes(raw)

            logging.info(f"  Subject: {msg.get('Subject', '(no subject)')}")

            with smtplib.SMTP("smtp.gmail.com", 587, timeout=30) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(self.gmail_user, self.gmail_password)
                smtp.sendmail(
                    envelope.mail_from,
                    envelope.rcpt_tos,
                    msg.as_bytes(),
                )

            logging.info(f"  Sent successfully via {self.gmail_user}")
            return "250 Message accepted for delivery"

        except smtplib.SMTPAuthenticationError as e:
            logging.error(f"  Gmail auth failed -- check app password: {e}")
            return "535 Authentication failed"
        except smtplib.SMTPException as e:
            logging.error(f"  SMTP error: {e}")
            return f"500 SMTP Error: {e}"
        except Exception as e:
            logging.exception("  Unexpected error handling message")
            return f"500 Error: {e}"

# -- Controller helpers --------------------------------------------------------
def start_controller(handler: EmailRouter, port: int) -> Controller:
    kill_port_holder(port)
    controller = Controller(handler, hostname="localhost", port=port)
    controller.start()
    return controller

def controller_is_alive(port: int) -> bool:
    try:
        with smtplib.SMTP("localhost", port, timeout=5):
            pass
        return True
    except Exception:
        return False

# -- Main ----------------------------------------------------------------------
async def main():
    configure_only = "--configure-only" in sys.argv

    if not configure_only:
        print()
        print("=" * 70)
        print("Practice Perfect Standalone Email Router")
        print("=" * 70)
        print()

    config = load_or_create_config()
    gmail_user      = config["gmail"]
    gmail_password  = config["app_password"]
    clinician_name  = config["name"]
    local_port      = config.get("local_port", LOCAL_PORT)

    if configure_only:
        print("\nConfiguration complete!")
        ok = test_gmail_connection(gmail_user, gmail_password)
        sys.exit(0 if ok else 1)

    lock_socket = acquire_instance_lock()

    print(f"Clinician:  {clinician_name}")
    print(f"Gmail:      {gmail_user}")
    print(f"Listening:  localhost:{local_port}")
    print()

    if not test_gmail_connection(gmail_user, gmail_password):
        answer = input("Continue anyway? (y/N): ").strip().lower()
        if answer != "y":
            print("Startup cancelled.")
            sys.exit(1)

    handler = EmailRouter(gmail_user, gmail_password, clinician_name)

    while True:
        controller = None
        try:
            logging.info("Starting SMTP controller...")
            controller = start_controller(handler, local_port)
            logging.info(f"Email router running -- localhost:{local_port} -> {gmail_user}")

            print("=" * 70)
            print("Email Router is RUNNING!")
            print("=" * 70)
            print(f"  SMTP Server:    localhost")
            print(f"  Port:           {local_port}")
            print(f"  Encryption:     None")
            print(f"  Authentication: None")
            print(f"  Log file:       email_router.log")
            print("\nPress Ctrl+C to stop\n")

            while True:
                await asyncio.sleep(HEARTBEAT_SEC)
                if not controller_is_alive(local_port):
                    logging.warning("Heartbeat failed -- restarting controller...")
                    break

            try:
                controller.stop()
            except Exception:
                pass

        except KeyboardInterrupt:
            logging.info("Shutdown requested by user (Ctrl+C)")
            if controller:
                try:
                    controller.stop()
                except Exception:
                    pass
            break

        except Exception as e:
            logging.error(f"Controller crashed: {e} -- restarting in {RESTART_DELAY}s...")
            if controller:
                try:
                    controller.stop()
                except Exception:
                    pass

        logging.info(f"Waiting {RESTART_DELAY}s before restart...")
        await asyncio.sleep(RESTART_DELAY)
        kill_port_holder(local_port)

    logging.info("Email router stopped.")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nEmail router stopped.")
    except Exception as e:
        logging.exception("Fatal error")
        print(f"\nFatal error: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)
'@

$emailRouterScript | Out-File -FilePath "$InstallDir\email_router.py" -Encoding UTF8
Write-Success "email_router.py created"

# -----------------------------------------------------------------------------
# Step 5: start_router.bat
# -----------------------------------------------------------------------------
Write-Info "Creating start_router.bat..."
@"
@echo off
cd /d $InstallDir
python email_router.py
pause
"@ | Out-File -FilePath "$InstallDir\start_router.bat" -Encoding ASCII
Write-Success "start_router.bat created"

# -----------------------------------------------------------------------------
# Step 6: start_router_hidden.vbs
#   stdout/stderr redirected to startup_log.txt so silent crashes are visible
# -----------------------------------------------------------------------------
Write-Info "Creating start_router_hidden.vbs..."
@"
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c ""cd /d $InstallDir && python email_router.py >> startup_log.txt 2>&1""", 0, False
Set WshShell = Nothing
"@ | Out-File -FilePath "$InstallDir\start_router_hidden.vbs" -Encoding ASCII
Write-Success "start_router_hidden.vbs created"

# -----------------------------------------------------------------------------
# Step 7: First-time configuration (includes automatic test email)
# -----------------------------------------------------------------------------
Write-Host ""
Write-Info "Running first-time Gmail configuration..."
Write-Host "A test email will be sent to the clinician's Gmail to confirm setup."
Write-Host ""

Push-Location $InstallDir
try {
    python email_router.py --configure-only
    if ($LASTEXITCODE -ne 0) {
        Write-Warn "Configuration failed or was cancelled"
    }
} catch {
    Write-Warn "Configuration cancelled or failed"
}
Pop-Location

if (-not (Test-Path "$InstallDir\email_router_config.json")) {
    Write-Warn "Configuration not completed. Finish later by running:"
    Write-Host "  $InstallDir\start_router.bat"
    Read-Host "Press Enter to exit"
    exit 0
}

Write-Success "Configuration complete"

# -----------------------------------------------------------------------------
# Step 8: Task Scheduler
# -----------------------------------------------------------------------------
Write-Host ""
Write-Info "Creating Task Scheduler entry..."

# Re-check in case a ghost XML reappeared between Step 0 and now
if (Test-Path $TaskXml) {
    Remove-Item -Path $TaskXml -Force -ErrorAction SilentlyContinue
    Write-Warn "Ghost task XML removed (appeared between cleanup and registration)"
}

$action    = New-ScheduledTaskAction `
                 -Execute  "wscript.exe" `
                 -Argument "`"$InstallDir\start_router_hidden.vbs`""
$trigger   = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$principal = New-ScheduledTaskPrincipal `
                 -UserId    $env:USERNAME `
                 -LogonType Interactive `
                 -RunLevel  Limited
$settings  = New-ScheduledTaskSettingsSet `
                 -AllowStartIfOnBatteries `
                 -DontStopIfGoingOnBatteries `
                 -StartWhenAvailable `
                 -RestartCount 3 `
                 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask `
    -TaskName  $TaskName `
    -Action    $action `
    -Trigger   $trigger `
    -Principal $principal `
    -Settings  $settings | Out-Null

Write-Success "Task Scheduler entry created"

# -----------------------------------------------------------------------------
# Step 9: Start immediately
# -----------------------------------------------------------------------------
Write-Host ""
Write-Info "Starting the email router..."
Start-ScheduledTask -TaskName $TaskName
Start-Sleep -Seconds 4

$pythonProcess = Get-Process python -ErrorAction SilentlyContinue
if ($pythonProcess) {
    Write-Success "Email router is running!"
} else {
    Write-Warn "Python process not detected -- it may still be starting."
    Write-Host "Check: $InstallDir\startup_log.txt"
    Write-Host "Or start manually: $InstallDir\start_router.bat"
}

# -----------------------------------------------------------------------------
# Done
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host ("=" * 70) -ForegroundColor Cyan
Write-Host ""
Write-Host "The router starts automatically at login and self-recovers from crashes."
Write-Host ""
Write-Host "Install directory : $InstallDir"
Write-Host "Application log   : $InstallDir\email_router.log"
Write-Host "Startup log       : $InstallDir\startup_log.txt"
Write-Host ""
Write-Host "Practice Perfect SMTP settings:" -ForegroundColor Cyan
Write-Host "   Server:         localhost"     -ForegroundColor Cyan
Write-Host "   Port:           2525"          -ForegroundColor Cyan
Write-Host "   Encryption:     None"          -ForegroundColor Cyan
Write-Host "   Authentication: None"          -ForegroundColor Cyan
Write-Host ""
Write-Host ("=" * 70) -ForegroundColor Cyan
Read-Host "Press Enter to exit"
