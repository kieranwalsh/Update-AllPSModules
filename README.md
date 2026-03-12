# Update-InstalledModule

Updates all locally installed PowerShell modules to the latest versions available on PSGallery.

## Features

- Automatically updates `PackageManagement` and `PowerShellGet` prerequisites before processing other modules.
- Installs to **CurrentUser** scope by default — no admin required.
- Optional `-AllUsers` switch for system-wide installs (requires elevation).
- Displays a formatted list of all modules with current vs. available versions.
- Falls back to uninstall/reinstall if a standard update fails.
- Excludes the `Az` and `Microsoft.Graph` meta-modules (to avoid reinstalling all submodules). Individual submodules (e.g., `Az.Accounts`, `Microsoft.Graph.Users`) are updated individually.

## Installation

```powershell
Install-Module -Name 'Update-InstalledModule'
```

## Usage

```powershell
# Update all modules (CurrentUser scope)
Update-InstalledModule

# Update all modules (AllUsers scope, requires admin)
Update-InstalledModule -AllUsers
```

## Sample Output

![Image of Update-InstalledModule sample](https://github.com/kieranwalsh/img/blob/main/Update-AllPSModules%20Sample.png)

![Gif of Update-InstalledModule in action](https://github.com/kieranwalsh/img/blob/main/Update-AllPSModules.gif)
