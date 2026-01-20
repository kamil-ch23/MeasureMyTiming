# MeasureMyTiming

A simple PowerShell-based project time tracking tool. Track time spent on multiple projects directly from your terminal.

## Features

- **Track Multiple Projects** - Time multiple projects independently
- **Cumulative Time Display** - See total time spent per project
- **Complete Projects** - Mark projects as done and archive them
- **Remove Projects** - Delete projects you no longer need
- **Automatic Backups** - Creates a backup before any data modification
- **Desktop Shortcut** - Quick access from your desktop

## Requirements

- Windows PowerShell 5.1 or later
- [ImportExcel](https://github.com/dfinke/ImportExcel) module

## Installation

1. Download `Setup-MeasureMyTiming.ps1`
2. Run the setup script:

```powershell
.\Setup-MeasureMyTiming.ps1
```

3. Follow the prompts to configure:
   - Installation path (default: `C:\MeasureMyTiming`)
   - Project name (default: `MeasureMyTiming`)

The setup will:
- Install the ImportExcel module if not present
- Create the project folder
- Create the Excel data file
- Generate the main script
- Create a desktop shortcut

### Advanced Installation

You can also run setup with parameters:

```powershell
# Custom path and name
.\Setup-MeasureMyTiming.ps1 -InstallPath "D:\MyTools\TimeTracker" -ProjectName "MyTimeTracker"

# Skip desktop shortcut creation
.\Setup-MeasureMyTiming.ps1 -SkipShortcut
```

## Usage

Launch the tool by double-clicking the desktop shortcut or running:

```powershell
powershell -File "C:\MeasureMyTiming\MeasureMyTiming.ps1"
```

### Menu Options

```
========== MeasureMyTiming ==========

Select a project to start timing:

1. ProjectA (02:30:45) [1]
2. ProjectB (01:15:30) [2]

P. Add Project [P]
C. Complete Project [C]
R. Remove Project [R]
X. Exit [X]

----- Completed -----
OldProject (05:42:18) [2026-01-15]
```

| Key | Action |
|-----|--------|
| `1-9` | Start timing the selected project |
| `P` | Add a new project and start timing |
| `C` | Mark a project as completed |
| `R` | Remove a project and all its records |
| `X` | Exit the application |

### Timing a Project

1. Select a project number or press `P` to add new
2. The timer starts immediately
3. Press `Enter` to stop timing
4. Time is automatically recorded

### Completing a Project

1. Press `C`
2. Select the project to complete
3. Project moves to the "Completed" section with total time

### Removing a Project

1. Press `R`
2. Select the project to remove
3. Confirm with `Y`
4. All records for that project are deleted

## File Structure

```
C:\MeasureMyTiming\
├── MeasureMyTiming.ps1      # Main script
├── MeasureMyTiming.xlsx     # Data file
└── ArchivedTiming\          # Backup folder
    ├── MeasureMyTiming-20012026-14-30-45.xlsx
    └── MeasureMyTiming-20012026-15-45-22.xlsx
```

## Data Storage

All timing data is stored in an Excel file with the following worksheets:

### Timing Sheet
| Column | Description |
|--------|-------------|
| Project_Name | Name of the project |
| Start_Time | When timing started |
| Stop_Time | When timing stopped |
| Overall_Time | Duration (hh:mm:ss) |

### Completed Sheet
| Column | Description |
|--------|-------------|
| Project_Name | Name of the completed project |
| Total_Time | Total cumulative time |
| Completed_Date | Date marked as complete |

## Backups

Automatic backups are created in the `ArchivedTiming` folder before any data modification:
- Adding time to a project
- Completing a project
- Removing a project

Backup filename format: `ProjectName-DDMMYYYY-HH-mm-ss.xlsx`

## Troubleshooting

### ImportExcel Module Not Found

Install manually:
```powershell
Install-Module -Name ImportExcel -Scope CurrentUser
```

### Permission Denied

Run PowerShell as Administrator or choose a different installation path.

### Excel File Locked

Close any Excel instances that have the data file open.

## License

MIT License - See [LICENSE](LICENSE) for details.

## Author

Kamil Chrabonszcz

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
