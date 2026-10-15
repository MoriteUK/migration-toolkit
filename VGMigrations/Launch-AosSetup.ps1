# Wrapper script to launch AOS Setup wizard
# This script loads the AOS Setup form and displays it

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load lib.ps1 for shared functions (Read-SharedConfig, Update-SharedConfig, etc.)
$libPath = Join-Path $scriptDir "lib.ps1"
if (Test-Path $libPath) {
    . $libPath
} else {
    Write-Error "Required file lib.ps1 not found in $scriptDir"
    exit 1
}

# Dot-source the aossetup.ps1 to load the Show-AosSetupForm function
. (Join-Path $scriptDir "aossetup.ps1")

# Launch the AOS Setup form
Show-AosSetupForm
