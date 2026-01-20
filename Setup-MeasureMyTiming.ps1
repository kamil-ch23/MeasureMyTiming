<#
.SYNOPSIS
    MeasureMyTiming Setup Script - A simple project time tracking tool for PowerShell.

.DESCRIPTION
    This setup script creates the MeasureMyTiming environment including:
    - Project folder structure
    - Excel file for time tracking data
    - Main PowerShell script (MeasureMyTiming.ps1)
    - Desktop shortcut for easy access
    
    MeasureMyTiming allows you to:
    - Track time spent on multiple projects
    - View cumulative time per project
    - Mark projects as completed
    - Remove projects
    - Automatic backup before any data changes

.EXAMPLE
    Run the setup with default settings (interactive prompts):
    
    PS> .\Setup-MeasureMyTiming.ps1

.EXAMPLE
    Run the setup with custom path and name:
    
    PS> .\Setup-MeasureMyTiming.ps1 -InstallPath "D:\MyTools\TimeTracker" -ProjectName "MyTimeTracker"

.PARAMETER InstallPath
    The directory where MeasureMyTiming will be installed.
    Default: User will be prompted, or C:\MeasureMyTiming if left blank.

.PARAMETER ProjectName
    The name for the project (used for folder, files, and shortcut).
    Default: User will be prompted, or MeasureMyTiming if left blank.

.PARAMETER SkipShortcut
    If specified, skips creating the desktop shortcut.

.NOTES
    VERSION:        1.0.0
    AUTHOR:         Kamil Chrabonszcz
    CREATION DATE:  2026-01-20
    REPOSITORY:     https://github.com/kamil-ch23/MeasureMyTiming
    LICENSE:        MIT
    
    Requirements:
    - PowerShell 5.1 or later
    - ImportExcel module (will prompt to install if missing)
    
    The ImportExcel module can be installed with:
    Install-Module -Name ImportExcel -Scope CurrentUser

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$InstallPath,
    
    [Parameter(Mandatory = $false)]
    [string]$ProjectName,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipShortcut
)

# ============================================
# BANNER
# ============================================
function Show-Banner {
    Write-Host ""
    Write-Host "============================================" 
    Write-Host "       MeasureMyTiming Setup Wizard         "
    Write-Host "============================================"
    Write-Host "  A simple project time tracking tool       "
    Write-Host "  Author: Kamil Chrabonszcz                 "
    Write-Host "  Version: 1.0.0                            "
    Write-Host "============================================"
    Write-Host ""
}

# ============================================
# CHECK PREREQUISITES
# ============================================
function Test-Prerequisites {
    Write-Host "[*] Checking prerequisites..."
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Host "[!] ERROR: PowerShell 5.1 or later is required."
        return $false
    }
    Write-Host "    PowerShell version: $($PSVersionTable.PSVersion) - OK"
    
    # Check ImportExcel module
    $importExcel = Get-Module -ListAvailable -Name ImportExcel
    if (-not $importExcel) {
        Write-Host ""
        Write-Host "[!] ImportExcel module is not installed."
        $install = Read-Host "    Would you like to install it now? (Y/N)"
        if ($install -eq 'Y' -or $install -eq 'y') {
            try {
                Write-Host "    Installing ImportExcel module..."
                Install-Module -Name ImportExcel -Scope CurrentUser -Force
                Write-Host "    ImportExcel module installed successfully."
            }
            catch {
                Write-Host "[!] ERROR: Failed to install ImportExcel module."
                Write-Host "    Please run: Install-Module -Name ImportExcel -Scope CurrentUser"
                return $false
            }
        }
        else {
            Write-Host "[!] ImportExcel module is required. Setup cannot continue."
            return $false
        }
    }
    else {
        Write-Host "    ImportExcel module: Installed - OK"
    }
    
    Write-Host ""
    return $true
}

# ============================================
# GET USER INPUT
# ============================================
function Get-UserInput {
    param(
        [string]$CurrentInstallPath,
        [string]$CurrentProjectName
    )
    
    $result = @{
        InstallPath = $CurrentInstallPath
        ProjectName = $CurrentProjectName
    }
    
    # Get Install Path
    if (-not $result.InstallPath) {
        Write-Host "[?] Enter installation path"
        Write-Host "    (Press Enter for default: C:\MeasureMyTiming)"
        $inputPath = Read-Host "    Path"
        if ([string]::IsNullOrWhiteSpace($inputPath)) {
            $result.InstallPath = "C:\MeasureMyTiming"
        }
        else {
            $result.InstallPath = $inputPath.Trim()
        }
    }
    
    # Get Project Name
    if (-not $result.ProjectName) {
        Write-Host ""
        Write-Host "[?] Enter project name (used for files and shortcut)"
        Write-Host "    (Press Enter for default: MeasureMyTiming)"
        $inputName = Read-Host "    Name"
        if ([string]::IsNullOrWhiteSpace($inputName)) {
            $result.ProjectName = "MeasureMyTiming"
        }
        else {
            $result.ProjectName = $inputName.Trim()
        }
    }
    
    Write-Host ""
    Write-Host "[*] Configuration Summary:"
    Write-Host "    Install Path: $($result.InstallPath)"
    Write-Host "    Project Name: $($result.ProjectName)"
    Write-Host "    Excel File:   $($result.ProjectName).xlsx"
    Write-Host "    Script File:  $($result.ProjectName).ps1"
    Write-Host ""
    
    $confirm = Read-Host "[?] Proceed with installation? (Y/N)"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "[!] Setup cancelled by user."
        return $null
    }
    
    return $result
}

# ============================================
# MAIN SCRIPT CONTENT
# ============================================
function Get-MainScriptContent {
    param(
        [string]$ExcelPath,
        [string]$ProjectName
    )
    
    $scriptContent = @"
<#
.SYNOPSIS
    $ProjectName - A simple project time tracking tool.

.DESCRIPTION
    Track time spent on multiple projects with features including:
    - Start/stop timing for projects
    - View cumulative time per project
    - Mark projects as completed
    - Remove projects
    - Automatic backup before any data changes

.EXAMPLE
    Run via desktop shortcut or directly:
    
    PS> .\$ProjectName.ps1

.NOTES
    VERSION:        1.0.0
    AUTHOR:         Kamil Chrabonszcz
    REPOSITORY:     https://github.com/kamil-ch23/MeasureMyTiming
    
    Requirements:
    - ImportExcel module
      Install with: Install-Module -Name ImportExcel -Scope CurrentUser

#>

# Import required module
Import-Module ImportExcel

`$excelPath = '$ExcelPath'

# Function to create backup before any file update
function Backup-ExcelFile {
    `$excelFolder = Split-Path -Path `$excelPath -Parent
    `$archiveFolder = Join-Path -Path `$excelFolder -ChildPath 'ArchivedTiming'
    
    # Create archive folder if it doesn't exist
    if (-not (Test-Path -Path `$archiveFolder)) {
        New-Item -Path `$archiveFolder -ItemType Directory | Out-Null
    }
    
    # Create backup filename with timestamp
    `$timestamp = Get-Date -Format "ddMMyyyy-HH-mm-ss"
    `$backupFileName = "$ProjectName-`$timestamp.xlsx"
    `$backupPath = Join-Path -Path `$archiveFolder -ChildPath `$backupFileName
    
    # Copy the file
    Copy-Item -Path `$excelPath -Destination `$backupPath
    Write-Host "Backup created: `$backupFileName"
}

# Function to read projects and calculate cumulative times
function Get-Projects {
    try {
        `$data = Import-Excel -Path `$excelPath -WorksheetName 'Timing' -ErrorAction Stop
        if (`$null -eq `$data) {
            return @(), @{}
        }
        `$projects = @(`$data | Select-Object -Property Project_Name -Unique | Where-Object { `$_.Project_Name -ne `$null -and `$_.Project_Name -ne '' })
        `$projectTimes = @{}
        foreach (`$row in `$data) {
            if (`$row.Overall_Time -and `$row.Project_Name) {
                try {
                    `$ts = [TimeSpan]::Parse(`$row.Overall_Time)
                    if (`$projectTimes.ContainsKey(`$row.Project_Name)) {
                        `$projectTimes[`$row.Project_Name] += `$ts
                    } else {
                        `$projectTimes[`$row.Project_Name] = `$ts
                    }
                }
                catch {
                    # Skip invalid time entries
                }
            }
        }
        return `$projects, `$projectTimes
    }
    catch {
        return @(), @{}
    }
}

# Function to get completed projects
function Get-CompletedProjects {
    try {
        # Check if Completed worksheet exists
        `$excelPackage = Open-ExcelPackage -Path `$excelPath
        `$worksheetNames = `$excelPackage.Workbook.Worksheets | ForEach-Object { `$_.Name }
        Close-ExcelPackage `$excelPackage -NoSave
        
        if (`$worksheetNames -contains 'Completed') {
            `$completedData = Import-Excel -Path `$excelPath -WorksheetName 'Completed' -ErrorAction Stop
            
            # Return empty if no data
            if (`$null -eq `$completedData) {
                return @()
            }
            
            # Force into array
            `$completedArray = @()
            if (`$completedData -is [Array]) {
                `$completedArray = `$completedData
            } else {
                `$completedArray = @(,`$completedData)
            }
            
            # Filter and return
            `$result = @()
            foreach (`$item in `$completedArray) {
                if (`$null -ne `$item.Project_Name -and `$item.Project_Name -ne '') {
                    `$result += `$item
                }
            }
            return `$result
        }
        return @()
    }
    catch {
        return @()
    }
}

# Function to display the menu
function Display-Menu {
    param (`$projects, `$projectTimes)
    
    Clear-Host
    Write-Host ""
    Write-Host "========== $ProjectName =========="
    Write-Host ""
    Write-Host "Select a project to start timing:"
    Write-Host ""
    
    if (`$null -eq `$projects -or `$projects.Count -eq 0) {
        Write-Host "(No projects timed yet)"
    }
    else {
        `$i = 1
        foreach (`$proj in `$projects) {
            `$cumulative = `$projectTimes[`$proj.Project_Name]
            `$cumulativeStr = if (`$cumulative) { "{0:hh\:mm\:ss}" -f `$cumulative } else { "00:00:00" }
            Write-Host "`$i. `$(`$proj.Project_Name) (`$cumulativeStr) [`$i]"
            `$i++
        }
    }
    
    Write-Host ""
    Write-Host "P. Add Project [P]"
    Write-Host "C. Complete Project [C]"
    Write-Host "R. Remove Project [R]"
    Write-Host "X. Exit [X]"
    Write-Host ""
    
    # Display completed projects section
    `$completedProjects = Get-CompletedProjects
    if (`$null -ne `$completedProjects -and `$completedProjects.Count -gt 0) {
        Write-Host "----- Completed -----"
        foreach (`$completed in `$completedProjects) {
            `$projName = `$completed.Project_Name
            `$totalTime = `$completed.Total_Time
            `$completedDate = `$completed.Completed_Date
            
            # Format the date
            if (`$completedDate -is [DateTime]) {
                `$dateStr = `$completedDate.ToString("yyyy-MM-dd")
            } else {
                `$dateStr = `$completedDate
            }
            
            Write-Host "`$projName (`$totalTime) [`$dateStr]"
        }
        Write-Host ""
    }
}

# Function to start and stop timing for a project
function Start-Timing {
    param (`$project)
    
    # Create backup before updating
    Backup-ExcelFile
    
    `$startTime = Get-Date
    Write-Host ""
    Write-Host "Timing '`$project'. Press Enter to stop..."
    Read-Host | Out-Null
    `$stopTime = Get-Date
    `$overallTime = `$stopTime - `$startTime
    `$overallTimeStr = "{0:hh\:mm\:ss}" -f `$overallTime
    `$newRow = [PSCustomObject]@{
        Project_Name = `$project
        Start_Time   = `$startTime
        Stop_Time    = `$stopTime
        Overall_Time = `$overallTimeStr
    }
    `$newRow | Export-Excel -Path `$excelPath -WorksheetName 'Timing' -Append
    Write-Host "Recorded: `$overallTimeStr"
}

# Function to complete a project
function Complete-Project {
    param (`$projects, `$projectTimes)
    
    Write-Host ""
    Write-Host "Select a project to mark as completed:"
    `$i = 1
    foreach (`$proj in `$projects) {
        `$cumulative = `$projectTimes[`$proj.Project_Name]
        `$cumulativeStr = if (`$cumulative) { "{0:hh\:mm\:ss}" -f `$cumulative } else { "00:00:00" }
        Write-Host "`$i. `$(`$proj.Project_Name) (`$cumulativeStr)"
        `$i++
    }
    Write-Host "0. Cancel"
    
    `$choice = Read-Host "Enter your choice"
    
    if (`$choice -eq '0') {
        return
    }
    
    if (`$choice -match '^\d+`$') {
        `$index = [int]`$choice - 1
        if (`$index -ge 0 -and `$index -lt `$projects.Count) {
            `$selectedProject = `$projects[`$index].Project_Name
            `$totalTime = `$projectTimes[`$selectedProject]
            `$totalTimeStr = if (`$totalTime) { "{0:hh\:mm\:ss}" -f `$totalTime } else { "00:00:00" }
            `$completedDate = Get-Date -Format "yyyy-MM-dd"
            
            # Create backup before updating
            Backup-ExcelFile
            
            # Add to Completed worksheet
            `$completedRow = [PSCustomObject]@{
                Project_Name   = `$selectedProject
                Total_Time     = `$totalTimeStr
                Completed_Date = `$completedDate
            }
            `$completedRow | Export-Excel -Path `$excelPath -WorksheetName 'Completed' -Append
            
            # Remove from Timing worksheet
            `$data = Import-Excel -Path `$excelPath -WorksheetName 'Timing'
            `$remainingData = @(`$data | Where-Object { `$_.Project_Name -ne `$selectedProject })
            
            if (`$remainingData.Count -gt 0) {
                `$remainingData = @(`$remainingData | Where-Object { `$_.Project_Name -ne `$null -and `$_.Project_Name -ne '' })
            }
            
            if (`$remainingData.Count -gt 0) {
                `$remainingData | Export-Excel -Path `$excelPath -WorksheetName 'Timing' -ClearSheet
            } else {
                # Write empty row to keep headers
                `$emptyRow = [PSCustomObject]@{
                    Project_Name = ""
                    Start_Time   = ""
                    Stop_Time    = ""
                    Overall_Time = ""
                }
                `$emptyRow | Export-Excel -Path `$excelPath -WorksheetName 'Timing' -ClearSheet
            }
            
            Write-Host ""
            Write-Host "Project '`$selectedProject' marked as completed!"
            Write-Host "Total time: `$totalTimeStr | Completed: `$completedDate"
            Start-Sleep -Seconds 2
        }
        else {
            Write-Host "Invalid choice"
            Start-Sleep -Seconds 1
        }
    }
    else {
        Write-Host "Invalid choice"
        Start-Sleep -Seconds 1
    }
}

# Function to remove a project
function Remove-Project {
    param (`$projects, `$projectTimes)
    
    Write-Host ""
    Write-Host "Select a project to remove (this will delete all records):"
    `$i = 1
    foreach (`$proj in `$projects) {
        `$cumulative = `$projectTimes[`$proj.Project_Name]
        `$cumulativeStr = if (`$cumulative) { "{0:hh\:mm\:ss}" -f `$cumulative } else { "00:00:00" }
        Write-Host "`$i. `$(`$proj.Project_Name) (`$cumulativeStr)"
        `$i++
    }
    Write-Host "0. Cancel"
    
    `$choice = Read-Host "Enter your choice"
    
    if (`$choice -eq '0') {
        return
    }
    
    if (`$choice -match '^\d+`$') {
        `$index = [int]`$choice - 1
        if (`$index -ge 0 -and `$index -lt `$projects.Count) {
            `$selectedProject = `$projects[`$index].Project_Name
            
            # Confirm deletion
            `$confirm = Read-Host "Are you sure you want to remove '`$selectedProject' and all its records? (Y/N)"
            
            if (`$confirm -eq 'Y' -or `$confirm -eq 'y') {
                # Create backup before updating
                Backup-ExcelFile
                
                # Remove from Timing worksheet
                `$data = Import-Excel -Path `$excelPath -WorksheetName 'Timing'
                `$remainingData = @(`$data | Where-Object { `$_.Project_Name -ne `$selectedProject })
                
                if (`$remainingData.Count -gt 0) {
                    `$remainingData = @(`$remainingData | Where-Object { `$_.Project_Name -ne `$null -and `$_.Project_Name -ne '' })
                }
                
                if (`$remainingData.Count -gt 0) {
                    `$remainingData | Export-Excel -Path `$excelPath -WorksheetName 'Timing' -ClearSheet
                } else {
                    # Write empty row to keep headers
                    `$emptyRow = [PSCustomObject]@{
                        Project_Name = ""
                        Start_Time   = ""
                        Stop_Time    = ""
                        Overall_Time = ""
                    }
                    `$emptyRow | Export-Excel -Path `$excelPath -WorksheetName 'Timing' -ClearSheet
                }
                
                Write-Host ""
                Write-Host "Project '`$selectedProject' and all its records have been removed."
                Start-Sleep -Seconds 2
            }
            else {
                Write-Host "Removal cancelled."
                Start-Sleep -Seconds 1
            }
        }
        else {
            Write-Host "Invalid choice"
            Start-Sleep -Seconds 1
        }
    }
    else {
        Write-Host "Invalid choice"
        Start-Sleep -Seconds 1
    }
}

# Function to add a new project
function Add-Project {
    Write-Host ""
    Write-Host "Enter new project name (or 0 to cancel):"
    `$newProject = Read-Host "Project name"
    
    if (`$newProject -eq '0') {
        return
    }
    
    if (`$newProject -and `$newProject.Trim() -ne '') {
        Start-Timing -project `$newProject.Trim()
    }
    else {
        Write-Host "Project name cannot be empty."
        Start-Sleep -Seconds 1
    }
}

# ============================================
# MAIN LOOP
# ============================================
while (`$true) {
    `$projects, `$projectTimes = Get-Projects
    Display-Menu -projects `$projects -projectTimes `$projectTimes
    `$choice = Read-Host "Enter your choice"
    if (`$choice -eq 'X' -or `$choice -eq 'x') {
        break
    }
    elseif (`$choice -eq 'P' -or `$choice -eq 'p') {
        Add-Project
    }
    elseif (`$choice -eq 'C' -or `$choice -eq 'c') {
        if (`$null -ne `$projects -and `$projects.Count -gt 0) {
            Complete-Project -projects `$projects -projectTimes `$projectTimes
        }
        else {
            Write-Host "No projects to complete."
            Start-Sleep -Seconds 1
        }
    }
    elseif (`$choice -eq 'R' -or `$choice -eq 'r') {
        if (`$null -ne `$projects -and `$projects.Count -gt 0) {
            Remove-Project -projects `$projects -projectTimes `$projectTimes
        }
        else {
            Write-Host "No projects to remove."
            Start-Sleep -Seconds 1
        }
    }
    elseif (`$choice -match '^\d+`$') {
        `$index = [int]`$choice - 1
        if (`$null -ne `$projects -and `$index -ge 0 -and `$index -lt `$projects.Count) {
            `$selectedProject = `$projects[`$index].Project_Name
            Start-Timing -project `$selectedProject
        }
        else {
            Write-Host "Invalid choice"
            Start-Sleep -Seconds 1
        }
    }
    else {
        Write-Host "Invalid choice"
        Start-Sleep -Seconds 1
    }
}
"@

    return $scriptContent
}

# ============================================
# CREATE EXCEL FILE
# ============================================
function New-ExcelFile {
    param(
        [string]$Path,
        [string]$ProjectName
    )
    
    Write-Host "[*] Creating Excel file..."
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
        
        # Create initial data row (will be empty but establishes headers)
        $timingRow = [PSCustomObject]@{
            Project_Name = ""
            Start_Time   = ""
            Stop_Time    = ""
            Overall_Time = ""
        }
        
        # Create Timing sheet
        $timingRow | Export-Excel -Path $Path -WorksheetName 'Timing' -AutoSize
        
        # Add Sheet2 and Sheet3 (required for proper Excel structure)
        $emptyData = [PSCustomObject]@{ Column1 = "" }
        $emptyData | Export-Excel -Path $Path -WorksheetName 'Sheet2' -Append
        $emptyData | Export-Excel -Path $Path -WorksheetName 'Sheet3' -Append
        
        Write-Host "    Excel file created successfully."
        return $true
    }
    catch {
        Write-Host "[!] ERROR: Failed to create Excel file."
        Write-Host "    Error: $_"
        return $false
    }
}

# ============================================
# CREATE DESKTOP SHORTCUT
# ============================================
function New-DesktopShortcut {
    param(
        [string]$ScriptPath,
        [string]$ShortcutName
    )
    
    Write-Host "[*] Creating desktop shortcut..."
    
    try {
        $shell = New-Object -ComObject WScript.Shell
        $desktop = [System.Environment]::GetFolderPath('Desktop')
        $shortcutPath = Join-Path $desktop "$ShortcutName.lnk"
        $shortcut = $shell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = 'C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe'
        $shortcut.Arguments = "-File `"$ScriptPath`""
        $shortcut.WorkingDirectory = Split-Path -Path $ScriptPath -Parent
        $shortcut.IconLocation = '%SystemRoot%\System32\SHELL32.dll,23'
        $shortcut.Save()
        
        # Release COM object
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        
        Write-Host "    Shortcut created: $shortcutPath"
    }
    catch {
        Write-Host "[!] WARNING: Failed to create desktop shortcut."
        Write-Host "    Error: $_"
    }
}

# ============================================
# MAIN SETUP EXECUTION
# ============================================

# Show banner
Show-Banner

# Check prerequisites
if (-not (Test-Prerequisites)) {
    Write-Host ""
    Write-Host "Setup failed. Please resolve the issues above and try again."
    Read-Host "Press Enter to exit"
    exit 1
}

# Get user input
$config = Get-UserInput -CurrentInstallPath $InstallPath -CurrentProjectName $ProjectName
if ($null -eq $config) {
    exit 0
}

$installPath = $config.InstallPath
$projectName = $config.ProjectName
$excelPath = Join-Path $installPath "$projectName.xlsx"
$scriptPath = Join-Path $installPath "$projectName.ps1"

Write-Host ""
Write-Host "[*] Starting installation..."

# Create directory if it doesn't exist
if (-not (Test-Path $installPath)) {
    Write-Host "[*] Creating directory: $installPath"
    New-Item -Path $installPath -ItemType Directory | Out-Null
}

# Create Excel file if it doesn't exist
if (-not (Test-Path $excelPath)) {
    if (-not (New-ExcelFile -Path $excelPath -ProjectName $projectName)) {
        Write-Host ""
        Write-Host "Setup failed. Could not create Excel file."
        Read-Host "Press Enter to exit"
        exit 1
    }
}
else {
    Write-Host "[*] Excel file already exists: $excelPath"
}

# Create main script
Write-Host "[*] Creating main script..."
$mainScriptContent = Get-MainScriptContent -ExcelPath $excelPath -ProjectName $projectName
$mainScriptContent | Out-File -FilePath $scriptPath -Encoding UTF8
Write-Host "    Script created: $scriptPath"

# Create desktop shortcut
if (-not $SkipShortcut) {
    New-DesktopShortcut -ScriptPath $scriptPath -ShortcutName $projectName
}

# Done
Write-Host ""
Write-Host "============================================"
Write-Host "    Setup Complete!"
Write-Host "============================================"
Write-Host ""
Write-Host "Installation Summary:"
Write-Host "  Install Path:  $installPath"
Write-Host "  Excel File:    $excelPath"
Write-Host "  Script File:   $scriptPath"
if (-not $SkipShortcut) {
    Write-Host "  Shortcut:      Desktop\$projectName.lnk"
}
Write-Host ""
Write-Host "To start tracking time:"
Write-Host "  - Double-click the '$projectName' shortcut on your desktop"
Write-Host "  - Or run: powershell -File `"$scriptPath`""
Write-Host ""
Write-Host "Backups are automatically created in:"
Write-Host "  $installPath\ArchivedTiming\"
Write-Host ""
Read-Host "Press Enter to exit"
