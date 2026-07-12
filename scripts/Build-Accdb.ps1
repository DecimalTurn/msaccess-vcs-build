#
# build accdb from source
#
param(
    [string]$SourceDir,
	[string]$TargetDir = "",
	[string]$FileName = "", # empty = name from vcs options
    [string]$VcsAddInPath = "" # empty = use default path (installed version)
   
)

# Helper: take a screenshot for diagnostics
function Take-Screenshot {
    param([string]$Label)
    $screenshotDir = Join-Path $curDir "screenshots"
    if (-not (Test-Path $screenshotDir)) { New-Item -Path $screenshotDir -ItemType Directory -Force | Out-Null }
    $ts = (Get-Date -Format "HHmmss")
    $path = Join-Path $screenshotDir "build_${Label}_${ts}.png"
    try {
        Add-Type -AssemblyName System.Windows.Forms 2>$null
        Add-Type -AssemblyName System.Drawing 2>$null
        $screen = [System.Windows.Forms.Screen]::PrimaryScreen
        $bounds = $screen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($bounds.X, $bounds.Y, 0, 0, $bounds.Size)
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $bitmap.Dispose()
        Write-Host "  [SCREENSHOT] $Label -> $path"
    } catch {
        Write-Host "  [SCREENSHOT] Failed: $_"
    }
}

# Check if the script is running under a Windows service account (SYSTEM, NETWORK SERVICE, LOCAL SERVICE)
$serviceAccounts = @('SYSTEM', 'NETWORK SERVICE', 'LOCAL SERVICE')
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
if ($serviceAccounts | Where-Object { $currentUser -match $_ }) {
    Write-Warning "Warning: This script is running under a Windows service account ($currentUser). Microsoft Access should not be executed as a service!"
}
else {
    Write-Host "Running script as user: $currentUser"
}

[string]$tempFileName = "VcsBuildTempApp"
[string]$accdbFileName = $tempFileName
if ($FileName -gt "") {
    $accdbFileName = $FileName
}

$curDir = $(Get-Location)
$accdbPath = "$curDir\$accdbFileName.accdb"

# open/create access file
$access = New-Object -ComObject Access.Application
$access.Visible = $true
if (-not (Test-Path $accdbPath)) {    
    $access.NewCurrentDatabase($accdbPath)
} 
else {
	$access.OpenCurrentDatabase($accdbPath)
}




[string]$addInProcessPath = ""
if ($VcsAddInPath -gt "") {
    $addInProcessPath = [System.IO.Path]::ChangeExtension($VcsAddInPath, "").TrimEnd('.')   
}
else {
    $appdata = $env:APPDATA
    $addInFolder = Join-Path $appdata "MSAccessVCS"
    $addInProcessPath = Join-Path $addInFolder "Version Control"
}

$addInPattern = "$addInProcessPath.accd[ae]"

if (-not (Test-Path $addInPattern)) {
    Write-Host "msaccess-vcs add-in not found: $addInPattern"
    Write-Host "Please install msaccess-vcs add-in first."
    exit 1
}

if (
    -not ([System.IO.Path]::IsPathRooted($SourceDir)) -or
    ($SourceDir -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceDir = Join-Path -Path (Get-Location) -ChildPath $SourceDir.TrimStart('\','/','.')
}

Write-Host "Add-in path: $addInProcessPath"
Write-Host "Current path: $curDir"
Write-Host "Source: $SourceDir"
Write-Host "TargetDir: $TargetDir"
Write-Host ""

Write-Host "Start msaccess-vcs build " -NoNewline

# Diagnostic: capture pre-build state
Write-Host ""
Write-Host "  [DIAG] Pre-build state:"
Write-Host "    CurrentProject.Name: $($access.CurrentProject.Name)"
Write-Host "    CurrentProject.FullName: $($access.CurrentProject.FullName)"
Write-Host "    CurrentProject.Path: $($access.CurrentProject.Path)"
Write-Host "    Forms.Count: $($access.Forms.Count)"
try {
    Write-Host "    VBProjects.Count: $(@($access.VBE.VBProjects).Count)"
    foreach ($proj in $access.VBE.VBProjects) {
        Write-Host "      VBProject: $($proj.Name) -> $($proj.FileName)"
    }
} catch {
    Write-Host "    VBProjects: (could not access: $_)"
}

# Step 1: Set interaction mode
try {
    $access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
    Write-Host "  SetInteractionMode: OK"
} catch {
    Write-Host "  SetInteractionMode: ERROR - $_"
    Write-Host "  [DIAG] Forms.Count after error: $($access.Forms.Count)"
    throw
}

# Step 2: Trigger build
try {
    $null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")
    Write-Host "  HandleRibbonCommand: OK"
} catch {
    Write-Host "  HandleRibbonCommand: ERROR - $_"
    Write-Host "  [DIAG] Forms.Count after error: $($access.Forms.Count)"
    Take-Screenshot "after_cmd_error"
}

Write-Host "  [DIAG] Post-command state:"
Write-Host "    Forms.Count: $($access.Forms.Count)"
Write-Host "    CurrentProject.Name: $($access.CurrentProject.Name)"
Take-Screenshot "after_cmd"

# VCS Build close tempApp and reopen new accdb => check 2x for Forms.Count
# Pump Windows messages explicitly so SetTimer callbacks fire in COM automation
Add-Type -AssemblyName System.Windows.Forms 2>$null
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$lastFormCount = -1
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 60)) {
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.Application]::DoEvents()
    $fc = $access.Forms.Count
    if ($fc -ne $lastFormCount) {
        Write-Host ("`n  [DIAG] Forms.Count: $fc ({0:f0}s)" -f $stopwatch.Elapsed.TotalSeconds)
        $lastFormCount = $fc
    }
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
Write-Host ("`n  [DIAG] First poll done: Forms.Count=$($access.Forms.Count), elapsed={0:f1}s" -f $stopwatch.Elapsed.TotalSeconds)
if ($stopwatch.Elapsed.TotalSeconds -ge 59) { Take-Screenshot "poll1_timeout" }

Start-Sleep -Seconds 3
Take-Screenshot "before_poll2"
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$lastFormCount = -1
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 60)) {
    Start-Sleep -Milliseconds 500
    [System.Windows.Forms.Application]::DoEvents()
    $fc = $access.Forms.Count
    if ($fc -ne $lastFormCount) {
        Write-Host ("`n  [DIAG] Forms.Count: $fc ({0:f0}s)" -f $stopwatch.Elapsed.TotalSeconds)
        $lastFormCount = $fc
    }
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
Write-Host ("`n  [DIAG] Second poll done: Forms.Count=$($access.Forms.Count), elapsed={0:f1}s" -f $stopwatch.Elapsed.TotalSeconds)
if ($stopwatch.Elapsed.TotalSeconds -ge 59) { Take-Screenshot "poll2_timeout" }
Write-Host " completed"

# Diagnostic: capture post-build state
Write-Host "  [DIAG] Post-build state:"
Take-Screenshot "post_build"
Write-Host "    CurrentProject.Name: '$($access.CurrentProject.Name)'"
Write-Host "    CurrentProject.FullName: '$($access.CurrentProject.FullName)'"
Write-Host "    CurrentProject.Path: '$($access.CurrentProject.Path)'"
Write-Host "    Forms.Count: $($access.Forms.Count)"
try {
    $db = $access.CurrentDb
    Write-Host "    CurrentDb.Name: '$($db.Name)'"
    Write-Host "    CurrentDb.TableDefs.Count: $($db.TableDefs.Count)"
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($db)
} catch {
    Write-Host "    CurrentDb: (could not access: $_)"
}
try {
    Write-Host "    IsCurrentDatabase: $($access.Application.IsCurrentDatabase)"
    Write-Host "    Visible: $($access.Visible)"
} catch {
    Write-Host "    App state: (could not access: $_)"
}

# Check if VCS build actually completed (form stays open in v5)
$buildCompleted = $false
$logDirs = @((Join-Path $SourceDir "logs"), (Join-Path $curDir "logs"))
foreach ($logDir in $logDirs) {
    if (Test-Path $logDir) {
        $latestLog = Get-ChildItem $logDir -Filter "Build_*.log" -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestLog) {
            $logContent = Get-Content $latestLog.FullName -Raw -ErrorAction SilentlyContinue
            Write-Host "  [VCS LOG] $($latestLog.Name):"
            Write-Host "  ----------"
            Write-Host $logContent
            Write-Host "  ----------"
            if ($logContent -match "Done\. \(|TOTAL RUNTIME") {
                $buildCompleted = $true
                Write-Host "  [DIAG] Build completed successfully per VCS log!"
                # Force-close the form since AutoClose timer may not fire in COM automation
                try {
                    Write-Host "  [DIAG] Closing frmVCSMain..."
                    $access.DoCmd.Close(2, "frmVCSMain")  # acForm=2
                    Start-Sleep -Seconds 2
                } catch {
                    Write-Host "  [DIAG] Form close: $_"
                }
                # Re-read CurrentProject after closing form
                $builtFileName = $access.CurrentProject.Name
                $builtFilePath = $access.CurrentProject.FullName
                Write-Host "  [DIAG] After form close: CurrentProject.Name='$builtFileName'"
            }
            break
        }
    }
}
if (-not $buildCompleted) {
    Write-Host "  [DIAG] No completed VCS build log found"
    $builtFileName = $access.CurrentProject.Name
    $builtFilePath = $access.CurrentProject.FullName
}

Start-Sleep -Seconds 1
Write-Host "Close Access " -NoNewline
$access.Quit(2)
Write-Host "." -NoNewline
Start-Sleep -Seconds 1
Write-Host "." -NoNewline
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
Write-Host "." -NoNewline
[GC]::WaitForPendingFinalizers()
Write-Host " completed"
Write-Host ""

if ( ($builtFileName -gt "") -and ($builtFileName -ne "$tempFileName.accdb") ) {
	Write-Host "Built: $builtFileName ($builtFilePath)"
} else {
	Write-Host "Build failed"
    if ([string]::IsNullOrEmpty($builtFileName)) {
        Write-Host "   (builtFileName is empty)"
    }
    else {
        Write-Host "   $builtFileName"
    }
    if ([string]::IsNullOrEmpty($builtFilePath)) {
        Write-Host "   (builtFilePath is empty)"
    } 
    else {
	    Write-Host "  $builtFilePath"
    }
	exit 1
}


# copy file to TargetDir
if ([string]::IsNullOrEmpty($FileName)) {
    $FileName = $builtFileName
}

$targetFilePath = $builtFilePath
$builtFilePathDir = [System.IO.Path]::GetDirectoryName($builtFilePath)
if (($TargetDir -gt "") -and ($TargetDir -ne  $builtFilePathDir) ) {
	Write-Host "Copy accdb to $TargetDir"
	New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null
    Copy-Item -Path $builtFilePath -Destination "$TargetDir\$FileName"
	Write-Host ""
	$targetFilePath = "$TargetDir\$FileName"
} elseif ($FileName -ne $builtFileName) {
	Rename-Item -Path ".\$builtFileName" -NewName $FileName -Force
}

$tempFilePath = Join-Path -Path $curDir -ChildPath ([System.IO.Path]::ChangeExtension($tempFileName, "accdb"))
if (Test-Path $tempFilePath) {
    Remove-Item -Path $tempFilePath -Force  
}

return "$targetFilePath"
