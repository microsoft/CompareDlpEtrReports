function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
 
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Debug', 'Info','Warning','Error')]
        [string]$Severity = 'Debug',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [Object]$ErrorRecord = $null,

        [Switch]$WriteHost,

        [Switch]$Checkpoint,

        [string]$Fore
    )

    [string]$exceptionString = ""
    if (-not ($null -eq $ErrorRecord))
    {
        $exceptionString = $ErrorRecord.ScriptStackTrace + "--- CommandName:" + $ErrorRecord.Exception.CommandName
    }

    if ($WriteHost)
    {
        try{
            Write-Host $Message -ForegroundColor $Fore
        }
        catch{
            Write-Host $Message
        }
    }

    [string]$CheckpointValue = ""
    if ($Checkpoint)
    {
        $CheckpointValue = "Checkpoint"
    }

    $LogFilePath = [Logs]::LogFilePath

    [pscustomobject]@{
        Time = (Get-Date -f g)
        Message = $Message
        Severity = $Severity
        ExceptionTrace = $exceptionString
        Checkpoint = $CheckpointValue
    } | Export-Csv -Path $LogFilePath -Append -NoTypeInformation
}

Class Logs{

    static [string] $LogFilePath

    static [void] CreateLogFile()
    {
    	[string]$basePath = ".\Logs\"
    	if (-not(Test-Path $basePath))
    	{
      	    New-Item -Path $basePath -ItemType directory
    	}

    	$LogFileBasePath = ".\Logs\LogFile_"
    	$date = (Get-Date).ToString("dd-MM-yyyy")
    	$LogFileBasePath = $LogFileBasePath + $date + "_"
    	$logFileNo = 1
    	[string]$NewLogFilePath = ""

    	while($true)
    	{
            $NewLogFilePath = $LogFileBasePath + $logFileNo.ToString() + ".txt"
            if (-not(Test-Path $NewLogFilePath)) 
            {
              break
            }
            $logFileNo += 1
        }

	    [Logs]::LogFilePath = $NewLogFilePath
    }
}