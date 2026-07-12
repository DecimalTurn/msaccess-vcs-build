#
# build accdb from source
#
param(
    [string]$SourceDir,
	[string]$TargetDir = "",
	[string]$FileName = "", # empty = name from vcs options
    [string]$VcsAddInPath = "" # empty = use default path (installed version)
   
)

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

# Step 1: Set interaction mode
try {
    $access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
} catch {
    Write-Host "  SetInteractionMode: ERROR - $_"
    throw
}

# Step 2: Trigger build
try {
    $null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")
} catch {
    Write-Host "  HandleRibbonCommand: ERROR - $_"
}

# VCS Build closes tempApp and reopens new accdb.
# Poll until the build completes: CurrentProject changes from temp to the target database.
Write-Host "." -NoNewline
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$buildDone = $false
while (-not $buildDone -and $stopwatch.Elapsed.TotalSeconds -lt 60) {
    Start-Sleep -Seconds 2
    $cpn = $access.CurrentProject.Name
    Write-Host "." -NoNewline
    if ($cpn -and $cpn -ne "$tempFileName.accdb") {
        $buildDone = $true
    }
}
$stopwatch.Stop()
Write-Host " completed"

$builtFileName = $access.CurrentProject.Name
$builtFilePath = $access.CurrentProject.FullName

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
Write-Host "."

# Ensure Access is fully closed and lock file is released
$lockFile = [System.IO.Path]::ChangeExtension($builtFilePath, "laccdb")
if (Test-Path $lockFile) {
    Write-Host "  Removing lock file: $lockFile"
    Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
}
# Kill any lingering Access processes
Get-Process -Name "MSACCESS" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
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
