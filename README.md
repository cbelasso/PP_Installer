# Practice Perfect Email Router Installer

One-click installer that sets up the Practice Perfect email router on Windows.

## For Clinicians: Installation

1. **Download** `PracticePerfect-Installer.exe` from the latest release
2. **Double-click** the .exe file
3. **Click "Yes"** when Windows asks for admin permission
4. Follow the on-screen prompts to enter your Gmail credentials
5. A test email will be sent to your Gmail account to confirm everything works
6. Done! The router starts automatically at login

### Gmail App Password

You'll need a Gmail App Password (not your regular password):
1. Go to https://myaccount.google.com/apppasswords
2. Select "Mail" and "Other (Custom name)"
3. Name it "Practice Perfect"
4. Copy the 16-character password into the installer

### SMTP Settings for Practice Perfect

Once installation is complete, configure Practice Perfect with:
- **Server:** localhost
- **Port:** 2525
- **Encryption:** None
- **Authentication:** None

## For Developers: Building

### Prerequisites
- .NET 8 SDK (included automatically in GitHub Actions)
- Windows machine (for local builds)

### Local Build on Windows

```bash
dotnet build src/InstallerWrapper.csproj -c Release
dotnet publish src/InstallerWrapper.csproj -c Release -o bin/Release/publish
```

The `.exe` will be in `bin/Release/publish/PracticePerfect-Installer.exe`

### Automatic Builds (GitHub Actions)

Every push to `main` automatically:
1. Compiles the C# wrapper on Windows
2. Bundles the PowerShell script
3. Creates an artifact (downloadable .exe + script)

To download:
1. Go to **Actions** tab
2. Click the latest "Build Installer EXE" workflow
3. Click **Artifacts** → Download `PracticePerfect-Installer`

### Creating a Release

To create an official release:
```bash
git tag v1.0.0
git push origin v1.0.0
```

This automatically creates a GitHub Release with the compiled .exe

## Project Structure

```
├── src/
│   ├── Program.cs              # C# wrapper (elevates privileges, runs PS1)
│   └── InstallerWrapper.csproj # Project configuration
├── .github/workflows/
│   └── build.yml               # GitHub Actions workflow
├── proxy_installer.ps1         # Main PowerShell installer script
└── README.md                   # This file
```

## How It Works

1. **C# Wrapper** (`Program.cs`): 
   - Requests admin elevation
   - Locates the PowerShell script in the same directory
   - Executes it with full admin privileges

2. **PowerShell Script** (`proxy_installer.ps1`):
   - Installs Python if needed
   - Installs aiosmtpd package
   - Prompts for Gmail credentials
   - Sends a test email to confirm setup
   - Creates a Windows Task Scheduler entry for auto-start
   - Starts the email router

## Troubleshooting

**"Script not found":** Make sure `proxy_installer.ps1` is in the same folder as the .exe

**"Admin elevation failed":** Run the .exe as Administrator manually

**Gmail authentication error:** 
- Verify your Gmail App Password (not your regular password)
- Re-run the installer to enter new credentials

**Port 2525 in use:** The installer automatically cleans up conflicting processes

## Questions?

Contact the Practice Perfect team for support.
