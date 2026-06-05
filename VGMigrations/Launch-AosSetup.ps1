# Wrapper script to launch AOS Setup wizard
# This script loads the AOS Setup form and displays it

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Dot-source the aossetup.ps1 to load the function
. (Join-Path $scriptDir "aossetup.ps1")

# Load shared config functions if they exist
$sharedConfigPath = Join-Path $scriptDir "shared-config.ps1"
if (Test-Path $sharedConfigPath) {
    . $sharedConfigPath
}

# Launch the AOS Setup form
Show-AosSetupForm
