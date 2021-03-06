# Find test folder
$scriptPath = $MyInvocation.MyCommand.Path
$testFolder = Split-Path $scriptPath -Parent
cd $testFolder

# Import module
Import-Module ..\CompareDlpEtrReports.psd1 -Force

# Create a log folder and file
[Logs]::CreateLogFile()

# Generate all the resources
. ".\Resources.ps1"

# Invoke all tests
$testFiles = "UtilsTests", "AggregatorTests", "CmdletUtilsTests"
foreach ($testFile in $testFiles)
{
    Invoke-Pester "$testFolder\$testFile.ps1"
}

# Clean up
if (Test-Path "$testFolder\Reports")
{
    Remove-Item -Recurse -Force "$testFolder\Reports"
}

Remove-Item -Recurse -Force "$testFolder\Logs"