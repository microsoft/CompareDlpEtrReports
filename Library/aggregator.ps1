Class Aggregator
{
    static [System.Collections.Generic.Dictionary[String,Collections.Generic.List[CommonReport]]] GetSpecificReports($ruleNames, [string]$baseCommand, [ReportType]$reportType, [string]$messageIdProperty)
    {
        $commonReports = New-Object System.Collections.Generic.Dictionary"[String,Collections.Generic.List[CommonReport]]"
        $largeLimit = $false
        foreach ($ruleName in $ruleNames)
        {
            # Construct the command
            [string]$command = ($baseCommand + '"' + $ruleName + '"') 

            # Call the cmdlet to retrieve reports
            Write-Log -Message "Trying to retrieve $reportType for rule ${ruleName}"

            $CmdletReports = Invoke-Expression $command
            $CmdletReportsInitialCount = ($CmdletReports | Measure-Object).Count

            Write-Log -Message "Retrieved ${CmdletReportsInitialCount} Etr reports for rule ${ruleName}"

            # Page size case
            if (($CmdletReportsInitialCount -eq 1000) -and (-not $largeLimit))
            {
                $largeLimit = $true
                Write-Log "Your Dataset has reached a large limit. Please reduce the time frame to improve results." -WriteHost
            }

            # Group same message together
            if (($CmdletReports | Measure-Object).Count -gt 0)
            {
                [Collections.Generic.List[CommonReport]]$commonReportsForRule = [Aggregator]::FilterAndGroupReportsBasedOnRule($CmdletReports, $ruleName, $reportType, $messageIdProperty)
                $commonReports.Add($ruleName, $commonReportsForRule)
            }
        }
        
        Write-Log -Message "Returning common reports ($reportType) after aggregation" -Checkpoint
        return $commonReports
    }

    static [Collections.Generic.List[CommonReport]] FilterAndGroupReportsBasedOnRule([Object[]]$CmdletReports, [string]$ruleName, [ReportType]$reportType, [string]$messageIdProperty)
    {
        [string]$rulePropertyName = ""
        if ($reportType -eq [ReportType]::DlpReport)
        {
            $rulePropertyName = "DlpComplianceRule"
        }
        else
        {
            $rulePropertyName = "TransportRule"
        }

        Write-Log -Message "Filtering reports only with rule name: $ruleName"
        $finalReports = $CmdletReports | Where-Object -Property $rulePropertyName -eq $ruleName

        return [Aggregator]::GenerateCommonReports($finalReports, $reportType, $messageIdProperty) 
    }

    static [Collections.Generic.List[CommonReport]]GenerateCommonReports(
            [Object[]]$specificReports,
            [ReportType]$reportType,
            [string]$messageIdProperty)
    {
        Write-Log -Message "Trying to generate common reports ($reportType)"

        # Initialize new list of common reports
        $commonReports = [Collections.Generic.List[CommonReport]]::new()
        $commonReportsByMessageId = [Collections.Generic.Dictionary[String, CommonReport]]::new()

        # Process each report individually
        foreach ($specificReport in $specificReports) 
        {     
            if ($null -eq $specificReport)
            {
                continue
            }

            $reportAction = $specificReport.Action
            [Constants.CommonAction]$action = [Utils]::TryParseAction($reportAction)

            if ($action -eq [Constants.CommonAction]::None)
            {
                if (-not($null -eq $reportAction) -and ($reportAction.Length -gt 0))
                {
                    Write-Log -Message "Action $reportAction is not supported in $reportType, please report this to Microsoft" -WriteHost -Severity Warning
                    #continue;
                } 
            }
            
            # Check if this message has been processed earlier
            [string]$messageId = $specificReport | Select-Object -ExpandProperty $messageIdProperty

            if ($null -eq $messageId)
            {
                continue
            }

            Write-Log -Message "Generating Common Report with action $reportAction and messageId $messageId"

            if (-not $commonReportsByMessageId.ContainsKey($messageId))
            {
                Write-Log -Message "No previous match found, creating new common report"
                [CommonReport]$psCommonReport = [CommonReport]::new($specificReport, $reportType, $action)
                $commonReportsByMessageId.Add($messageId, $psCommonReport)
            }
            else
            {
                Write-Log -Message "Match found, adding to the actions"
                # Message processed earlier, add actions to existing object
                [CommonReport]$matchedReport = $commonReportsByMessageId[$messageId]
                $matchedReport.actions = $matchedReport.actions -bor $action
            }
        }

        # Convert dictionary to list
        $commonReportsByMessageId.Values.ForEach({ 
            if (($_ | Measure-Object).Count -eq 1) 
            { 
                $commonReports.Add($_) 
            } 
        })
         
        return $commonReports
    }
}