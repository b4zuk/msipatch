# MSIPatch

MSIPatch is a utility for injecting files or defining custom actions into existing Windows Installer (MSI) packages. It allows you to programmatically modify MSI tables and associated cabinet files to include additional content, such as DLLs, executables, or custom scripts.

## Features

- Inject a file into an MSI and update all relevant tables.
- Modify `CAB` files embedded in the MSI.
- Target specific system folders (like `System32`, `Desktop`, `Program Files`, etc.).
- Automatically regenerate necessary IDT tables (`File.idt`, `Component.idt`, `Directory.idt`, etc.).
- Support for both x86 and x64 MSI packages.

## Requirements

Ensure the following tools are installed and available in your PATH:

- `msidump`
- `msibuild`
- `cabextract`
- `gcab`

These are required to extract and rebuild MSI and CAB files.

## Installation

This tool does not require installation. Just clone the repository and run the script:

```bash
git clone https://github.com/yourusername/msipatch.git
cd msipatch
python msipatch.py --help
```

## Usage

### List all valid directory keys:

```bash
python msipatch.py --list
```

### Inject a file into an MSI:

```console
python msipatch.py \
  --msi path/to/original.msi \
  --file path/to/file.dll \
  --dest system32 \
  --name file.dll \
  --cab output.dll \
  --comp FileInjectionComponent \
  --arch x64
```

### Arguments

| Option         | Description                                                                 |
|----------------|-----------------------------------------------------------------------------|
| `--msi`, `-m`  | Path to the original MSI file                                               |
| `--file`, `-i` | File to inject into the MSI                                                 |
| `--dest`, `-d` | Destination folder key (e.g., `system32`, `desktop`, etc.)                 |
| `--name`, `-n` | Target filename for the dropped file                                        |
| `--cab`, `-c`  | Filename inside the CAB archive                                             |
| `--comp`, `-C` | Component name to use in the MSI tables                                     |
| `--arch`, `-a` | Architecture of the MSI (`x86` or `x64`)                                    |
| `--list`, `-l` | List available destination folder keys and exit                             |

> **Note:** When using `--file`, all of `--dest`, `--name`, `--cab`, and `--comp` are required.

## Folder Key Reference

You can see all valid destination folder keys using:

```bash
python msipatch.py --list
```

These correspond to special folder IDs used by Windows Installer (`System32`, `AppData`, `Program Files`, `Desktop`, etc.).

## Common Categories

### User-Specific Folders

| Destination     | MSI Folder ID       | Default Path                               |
|-----------------|---------------------|--------------------------------------------|
| desktop         | `DesktopFolder`     | `C:\Users\Username\Desktop`                |
| documents       | `PersonalFolder`    | `C:\Users\Username\Documents`              |
| downloads       | `DownloadsFolder`   | `C:\Users\Username\Downloads`              |
| pictures        | `MyPicturesFolder`  | `C:\Users\Username\Pictures`               |
| music           | `MyMusicFolder`     | `C:\Users\Username\Music`                  |
| videos          | `MyVideoFolder`     | `C:\Users\Username\Videos`                 |
| start menu      | `StartMenuFolder`   | `C:\Users\Username\AppData\Roaming\Microsoft\Windows\Start Menu` |
| programs        | `ProgramsFolder`    | Subfolder under Start Menu                 |
| startup         | `StartupFolder`     | Startup folder under Start Menu            |
| admin tools     | `AdminToolsFolder`  | User-specific administrative tools         |
| appdata         | `AppDataFolder`     | `C:\Users\Username\AppData\Roaming`        |
| localappdata    | `LocalAppDataFolder`| `C:\Users\Username\AppData\Local`          |

---

### Common (All Users) Folders

| Destination          | MSI Folder ID             | Default Path                                           |
|----------------------|---------------------------|--------------------------------------------------------|
| common desktop       | `CommonDesktopFolder`     | `C:\Users\Public\Desktop`                              |
| common programs      | `CommonProgramsFolder`    | `C:\ProgramData\Microsoft\Windows\Start Menu\Programs`|
| common start menu    | `CommonStartMenuFolder`   | `C:\ProgramData\Microsoft\Windows\Start Menu`         |
| common startup       | `CommonStartupFolder`     | `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup` |
| common admin tools   | `CommonAdminToolsFolder`  | `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Administrative Tools` |
| programdata          | `CommonAppDataFolder`     | `C:\ProgramData`                                       |

---

### System Folders

| Destination      | MSI Folder ID         | Default Path                  |
|------------------|-----------------------|-------------------------------|
| windows          | `WindowsFolder`       | `C:\Windows`                  |
| system32         | `System64Folder` \| `SystemFolder` | `C:\Windows\System32` |
| syswow64         | `SystemFolder` \| `System64Folder` | `C:\Windows\SysWOW64` |
| fonts            | `FontsFolder`         | `C:\Windows\Fonts`            |
| temp             | `TempFolder`          | `%TEMP%`                      |

---

### Program Files

| Destination             | MSI Folder ID             | Default Path                        |
|-------------------------|---------------------------|-------------------------------------|
| program files           | `ProgramFilesFolder` / `ProgramFiles64Folder` | `C:\Program Files` / `C:\Program Files (x86)` |
| program files (x86)     | `ProgramFilesFolder`       | `C:\Program Files (x86)`            |
| common files            | `CommonFilesFolder` / `CommonFiles64Folder` | Inside Program Files               |

---

## Disclaimer

This tool is intended **strictly for red teaming, security research, and educational purposes**. Unauthorized use of this software in production environments or against systems without explicit permission is **strictly prohibited**. Misuse of MSIPatch may violate laws or terms of service. Always operate responsibly and with consent.
