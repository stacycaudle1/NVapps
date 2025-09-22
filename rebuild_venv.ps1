<#
Rebuild Python virtual environment for NVapps.
Usage: Right-click -> Run with PowerShell (or run from an elevated terminal if needed).
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Host '=== NVapps Virtual Environment Rebuild ===' -ForegroundColor Cyan
function Stop-PythonProcesses {
    Write-Host 'Stopping running python processes (if any)...'
    Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Try-RenameVenv {
    param($Path)
    if (Test-Path $Path) {
        try {
            # Build a safe new name from the leaf component and keep directory semantics correct
            $leaf = Split-Path -Leaf $Path
            $newName = "${leaf}_old"
            Rename-Item -Path $Path -NewName $newName -ErrorAction Stop
            Write-Host "Renamed $Path to $newName" -ForegroundColor Yellow
            return ($Path + '_old')
        } catch {
            Write-Host 'Rename failed; will attempt robocopy mirror removal.' -ForegroundColor DarkYellow
            Write-Host "Rename error: $_" -ForegroundColor DarkYellow
            return $null
        }
    }
}
<#
Rebuild Python virtual environment for NVapps.
Usage: Right-click -> Run with PowerShell (or run from an elevated terminal if needed).
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'
Write-Host '=== NVapps Virtual Environment Rebuild ===' -ForegroundColor Cyan

function Stop-PythonProcesses {
    Write-Host 'Stopping running python processes (if any)...'
    Get-Process python -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

function Try-RenameVenv {
    param($Path)
    if (Test-Path $Path) {
        try {
            # Build a safe new name from the leaf component and keep directory semantics correct
            $leaf = Split-Path -Leaf $Path
            $newName = "${leaf}_old"
            Rename-Item -Path $Path -NewName $newName -ErrorAction Stop
            Write-Host "Renamed $Path to $newName" -ForegroundColor Yellow
            return ($Path + '_old')
        } catch {
            Write-Host 'Rename failed; will attempt robocopy mirror removal.' -ForegroundColor DarkYellow
            Write-Host "Rename error: $_" -ForegroundColor DarkYellow
            return $null
        }
    }
}

function Force-RemoveVenv {
    param($Path)
    if (-not (Test-Path $Path)) { return }
    Write-Host 'Removing existing venv via robocopy mirror trick...' -ForegroundColor Yellow
    $empty = '_empty_dir'
    if (-not (Test-Path $empty)) { New-Item -ItemType Directory -Path $empty | Out-Null }
    robocopy $empty $Path /MIR /NFL /NDL /NJH /NJS /NP > $null
    Remove-Item $empty -Recurse -Force -ErrorAction SilentlyContinue
    try { Remove-Item $Path -Recurse -Force -ErrorAction Stop } catch { Write-Host "Residual removal warning: $_" -ForegroundColor DarkYellow }
}

function Create-Venv {
    param($Interpreter)
    Write-Host "Creating new venv with: $Interpreter" -ForegroundColor Cyan
    # Support interpreters with arguments like 'py -3.12'
    $parts = @()
    if ($Interpreter -is [string]) {
        $parts = $Interpreter -split ' '
    } else {
        $parts = @($Interpreter)
    }
    if ($parts.Count -gt 1) {
        & $parts[0] $parts[1..($parts.Count-1)] -m venv .venv
    } else {
        & $parts[0] -m venv .venv
    }
    if (-not (Test-Path '.venv\Scripts\Activate.ps1')) { throw 'Venv creation failed (no Scripts/Activate.ps1)' }
}

function Activate-Venv {
    Write-Host 'Activating virtual environment...'
    # Dot-source the activation script so it modifies the current session environment
    . .\.venv\Scripts\Activate.ps1
}

function Install-Dependencies {
    Write-Host 'Upgrading pip & installing requirements...' -ForegroundColor Cyan
    python -m pip install --upgrade pip
    if (Test-Path 'requirements.txt') {
        python -m pip install -r requirements.txt
    } else {
        Write-Host 'requirements.txt not found; skipping dependency install.' -ForegroundColor Yellow
    }
}

function Verify-Environment {
    Write-Host 'Verifying key imports...' -ForegroundColor Cyan
    $pyCode = @'
import sys, importlib
mods = ["openpyxl","pandas","matplotlib","tkinter"]
print("Interpreter:", sys.executable)
print("Version:", sys.version.split()[0])
for m in mods:
    try:
        importlib.import_module(m)
        print("[OK]", m)
    except Exception as e:
        print("[FAIL]", m, "-", e)
'@
    & python -c $pyCode
}

# Main flow
Stop-PythonProcesses
$renamed = Try-RenameVenv -Path '.venv'
if (-not $renamed) { Force-RemoveVenv -Path '.venv' }

# Choose interpreter
$python = 'python'
try {
    $version = & $python -c 'import sys;print(sys.version_info[:])' 2>$null
    if (-not $version) { throw 'python not on PATH' }
} catch {
    Write-Host 'Base python not found via "python"; trying py launcher...' -ForegroundColor Yellow
    $python = 'py -3.12'
}

try {
    Create-Venv -Interpreter $python
    Activate-Venv
    Install-Dependencies
    Verify-Environment
    Write-Host 'Rebuild complete.' -ForegroundColor Green
    if ($renamed) { Write-Host "You can now delete leftover $renamed when satisfied." -ForegroundColor Yellow }
} catch {
    Write-Host "FAILED: $_" -ForegroundColor Red
    Write-Host 'Manual steps: remove .venv, ensure python runs, recreate venv.' -ForegroundColor Red
}
