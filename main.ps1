function Compare-DlpEtrReports {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ReportFolderName,
 
        [Parameter()]
        [DateTime]$StartDate = $null,

        [Parameter()]
        [DateTime]$EndDate = $null,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$AdminEmailAddress
    )
    try
    {
        # Creating log file which can be used throughout execution
	    [Logs]::CreateLogFile()
	    Write-Log -Message "Beginning execution..." -Severity Debug

        # Initialize the cmdlet parameters if not specified in the list
        if (-not($EndDate -and $StartDate))
        {
            $EndDate = Get-Date
            $StartDate = $EndDate.AddDays(-1)
        }
        
        [string]$EndDate = $EndDate.ToString('MM-dd-yyyy hh:mm:ss')
        [string]$StartDate = $StartDate.ToString('MM-dd-yyyy hh:mm:ss')

        Write-Log -Message "Initialized Cmdlet Parameters. StartDate: $StartDate, EndDate: $EndDate" -Severity Debug

        # Get OAuth token for cmdlet invocations
        [CmdletUtils]::GetOAuthTokenForCmdlets($AdminEmailAddress)
    
        # Retrieve all the (ETR) transport rules
        Write-Host "Retrieving Transport Rules..."
        [Object[]]$Rules = [CmdletUtils]::GetTransportRules()

        # Retrieve all DLP rules
        Write-Host "Retrieving DLP Rules..."
        $AllDlpRuleNames = [CmdletUtils]::GetDlpRulesNames() 
        
        # Rules with non-blocking actions
        [Object[]]$NonBlockingRules = [CmdletUtils]::GetNonBlockingTransportRules($Rules)
        $NonBlockingRuleNames = $NonBlockingRules | Select-Object -ExpandProperty "Name"

        # Rules with blocking actions
        $BlockingRules = $Rules | Where-Object -Property Name -NotIn $NonBlockingRuleNames
        $BlockingRulesByPolicy = [CmdletUtils]::GroupTransportRulesByPolicy($BlockingRules)
        
        $EtrReportCommand = "Get-MailDetailDlpPolicyReport -StartDate '$StartDate' -EndDate '$EndDate' -TransportRule "
        $Global:AllMessagesActedUponByBlockingRules = [CmdletUtils]::GetMessagesActedByBlockingRules($BlockingRules, $EtrReportCommand)

        # If no rules are present, end it
        if (($NonBlockingRuleNames | Measure-Object).Count -eq 0)
        {
            Write-Log "No transport rules found, ending execution..." -WriteHost -Severity Info
        }

        Write-Log -Message "Downloading and aggregating ETR reports..." -WriteHost -Checkpoint
        [System.Collections.Generic.Dictionary[String,Collections.Generic.List[CommonReport]]]$commonEtrReports = [Aggregator]::GetSpecificReports($NonBlockingRuleNames, $EtrReportCommand, [ReportType]::EtrReport, "messageID")
        
        $DlpReportCommand = "Get-DlpDetailReport -StartDate '$StartDate' -EndDate '$EndDate' -Source EXCH -DlpComplianceRule "
        Write-Log -Message "Downloading and aggregating DLP reports..." -WriteHost -Checkpoint
        [System.Collections.Generic.Dictionary[String,Collections.Generic.List[CommonReport]]]$commonDlpReports = [Aggregator]::GetSpecificReports($NonBlockingRuleNames, $DlpReportCommand, [ReportType]::DlpReport, "Location")

        ##################################################################################################################################

        Write-Log "Completed Aggregating, Analyzing..." -WriteHost -Checkpoint

        $soloDlpReports = [Collections.Generic.List[SoloReport]]::new()
        $soloEtrReports = [Collections.Generic.List[SoloReport]]::new()
        $etrDlpMatchedReports = [Collections.Generic.List[CombinedReport]]::new()
        $allDlpReports = [Collections.Generic.List[CommonReport]]::new()
        $allEtrReports = [Collections.Generic.List[CommonReport]]::new()
        $restrictedEtrReports = [Collections.Generic.List[SoloReport]]::new()

        # Analyze DLP reports
        [Analyzer]::AnalyzeDlpReports($commonDlpReports, $commonEtrReports, $allDlpReports, $etrDlpMatchedReports, $soloDlpReports, $Rules)

        # Analyze ETR reports
        [Analyzer]::AnalyzeEtrReports($commonDlpReports, $commonEtrReports, $allEtrReports, $soloEtrReports, $restrictedEtrReports, $AllDlpRuleNames)
           
        # Group each report by the policy
        $mainReports = [Utils]::GenerateMainReport($commonDlpReports, $commonEtrReports, $BlockingRulesByPolicy)

        Write-Host "----------------------------------------------------------------------------------"
    
        [Utils]::ExportReportsToExcel($ReportFolderName, $mainReports, $allDlpReports, $allEtrReports, $etrDlpMatchedReports, $soloDlpReports, $soloEtrReports, $restrictedEtrReports)
    }
    catch [System.Exception]
    {
        Write-Host "There's been an error during execution, please contact Microsoft. Please pass the logs while raising support ticket."
        Write-Log -Message $PSItem.ToString() -ErrorRecord $PSItem -Severity Error
    }
    finally
    {
        $LogPath = [Logs]::LogFilePath
        Write-Host "The logs for this execution is generated in the path: $LogPath"
    }
}
