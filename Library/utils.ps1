class Utils
{
    static [string]GetReportFolderPath([string]$ReportFolderName)
    {    
        [string]$basePath = ".\Reports\"
        if (-not(Test-Path $basePath))
        {
            New-Item -Path $basePath -ItemType directory
        }

        [string]$finalReportFolderName = ""
        if (-not($null -eq $ReportFolderName) -and ($ReportFolderName.Length -gt 0))
        {
            $reportFilePath = $basePath + $ReportFolderName
        }
        else
        {
            $dateNow = Get-Date
            $reportFilePath = $basePath + "Reports_" + $dateNow.ToString("dd-MM-yyyy")
        }

        $reportNumber = 1
        $baseReportFilePath = $reportFilePath + "_"

        while(Test-Path $reportFilePath)
        {
            $reportFilePath = $baseReportFilePath + $reportNumber.ToString()
            $reportNumber = $reportNumber + 1

        }

        New-Item -Path $reportFilePath -ItemType directory
             
        return $reportFilePath + "\"
    }

    static [Constants.CommonAction] GetActionsFromRule($Rule)
    {
        $actions = $Rule.Actions
        [string]$action = ""
        [Constants.CommonAction]$reportActions = [Constants.CommonAction]::None
        foreach ($action in $actions)
        {
            $action = $action.Replace("Microsoft.Exchange.MessagingPolicies.Rules.Tasks.", "");
            $action = $action.Replace("Action", "");
            $parsedAction = [Utils]::TryParseAction($action)
                
            if ($parsedAction -eq [Constants.CommonAction]::None)
            {
                Write-Log "Cannot parse action $action in GetActionsFromRule" -Severity Error
                continue
            }
            
            # If there is a notify + block, make sure we add the block action here
            if ($parsedAction -eq [Constants.CommonAction]::NotifySender)
            {
                $notificationType = $Rule.SenderNotificationType
                if (-not($notificationType -eq "NotifyOnly"))
                {
                    $parsedAction = $parsedAction -bor [Constants.CommonAction]::RejectMessage
                }
            }

            Write-Log -Message "Found action $parsedAction in GetActionsFromRule"
            $reportActions = $reportActions -bor $parsedAction
        }
            
        return $reportActions
    }
    
    static [Constants.CommonAction]TryParseAction([string]$reportAction)
    {
        # Try to parse appropriate action
        Write-Log -Message "Trying to parse action $reportAction in TryParseAction"
        [Constants.CommonAction]$action = [Constants.CommonAction]::None
        try
        {
            $action = [Enum]::Parse([Constants.CommonAction], $reportAction)
        }
        catch
        {
            try{
                if (-not($null -eq $reportAction) -and ($reportAction.Length -gt 0))
                {
                    $action = $Global:EtrActionToCommonActionMap[$reportAction]
                }
            }
            catch{
                Write-Log -Message "Could not parse action in TryParseActoin"
            }    
        }

        Write-Log -Message "Returning action $action from TryParseAction"
        return $action
    }

    static [void]GroupRuleReportsByPolicy($reports, $PolicyToRulesMapping)
    {
        foreach ($ruleName in $reports.Keys)
        {
            [CommonReport]$commonReport = $reports[$ruleName][0]
            
            [int]$commonReportCount = ($commonReport | Measure-Object).Count
            if (($commonReportCount -eq 0) -or ($null -eq $commonReport.policy))
            {
                Write-Log -Message "Cannot find policy name for $commonReportCount rules under rule: $ruleName" -Severity Error
                continue
            }
            
            [string]$policyName = $commonReport.policy
            
            if (-not $PolicyToRulesMapping.ContainsKey($policyName))
            {
                $PolicyToRulesMapping.Add($policyName, [Collections.Generic.List[String]]::new())
            }

            if (-not $PolicyToRulesMapping[$policyName].Contains($ruleName))
            {
                $PolicyToRulesMapping[$policyName].Add($ruleName)
            }
        }
    }

    static [Collections.Generic.List[MainReport]] GenerateMainReport($commonDlpReports, $commonEtrReports, $BlockingRulesByPolicy)
    {
        [MainReport]$ruleReport = $null
        $PolicyToRulesMapping = [Collections.Generic.Dictionary[string, Collections.Generic.List[String]]]::new()
        $mainReports = [Collections.Generic.List[MainReport]]::new()
        $mainReports.Add([MainReport]::new(""))

        [Utils]::GroupRuleReportsByPolicy($commonDlpReports, $PolicyToRulesMapping)
        [Utils]::GroupRuleReportsByPolicy($commonEtrReports, $PolicyToRulesMapping)

        foreach ($policyName in $PolicyToRulesMapping.Keys)
        {
            $mainReports.Add([MainReport]::new($policyName))
            foreach ($ruleName in $PolicyToRulesMapping[$policyName])
            {
                $DlpReportCount = ($commonDlpReports[$ruleName] | Measure-Object).Count
                $EtrReportCount = ($commonEtrReports[$ruleName] | Where-Object -Property tempReason -eq $null | Measure-Object).Count
                $EtrReportRestrictionCount = ($commonEtrReports[$ruleName] | Measure-Object).Count

                $mainReports.Add([MainReport]::new($ruleName, $DlpReportCount, $EtrReportCount, $EtrReportRestrictionCount))
            }
        }

        if ($null -ne $BlockingRulesByPolicy)
        {
            # Blocked Rule Names to be added to the end
            $mainReports.Add([MainReport]::new(""))
            $mainReports.Add([MainReport]::new("The following rules cannot be processed because of restrictive actions"))
            foreach ($BlockedPolicyName in $BlockingRulesByPolicy.Keys)
            {
                $mainReports.Add([MainReport]::new($BlockedPolicyName))
                foreach($BlockedRule in $BlockingRulesByPolicy[$BlockedPolicyName])
                {
                    $mainReports.Add([MainReport]::new("", $BlockedRule))
                }
            }
        }

        return $mainReports
    }

    static [void] ExportReportsToExcel($ReportFolderName, $mainReports, $allDlpReports, $allEtrReports, $etrDlpMatchedReports, $soloDlpReports, $soloEtrReports, $restrictedEtrReports)
    {
        # Exporting to excel 
        $reportFilePath = [Utils]::GetReportFolderPath($ReportFolderName)
	
	    if (-not($null -eq $mainReports))
	    {
            $mainReportFilePath = $reportFilePath + $Global:Strings.MainReportFilePath
            # -Headers $Global:Strings.HeadersMainReport
            $mainReports | Export-Csv -Path $mainReportFilePath -NoTypeInformation
	    }

	    if (-not($null -eq $allDlpReports))
	    { 
            $allDlpReportsFilePath = $reportFilePath + $Global:Strings.AllDLPReportsFilePath
            # -Headers $Global:Strings.HeadersAllDlpReports
            $allDlpReports | Export-Csv -Path $allDlpReportsFilePath -NoTypeInformation
	    }

	    if (-not($null -eq $allEtrReports))
	    {        
            $allEtrReportsFilePath = $reportFilePath + $Global:Strings.AllETRReportsFilePath
            #  -Headers $Global:Strings.HeadersAllEtrReports
	        $allEtrReports | Export-Csv -Path $allEtrReportsFilePath -NoTypeInformation
        }

	    if (-not($null -eq $etrDlpMatchedReports))
	    {
            $etrDlpMatchedReportsFilePath = $reportFilePath + $Global:Strings.ETRDLPParityReportsFilePath
            # -Headers $Global:Strings.HeadersMatchedReports
	        $etrDlpMatchedReports | Export-Csv -Path $etrDlpMatchedReportsFilePath -NoTypeInformation
        }
	
	    if (($soloDlpReports | Measure-Object).Count -gt 0)
	    {
            $soloDlpReportsFilePath = $reportFilePath + $Global:Strings.SoloDLPReportsFilePath
            # -Headers $Global:Strings.HeadersSoloDlpReports
	        $soloDlpReports | Export-Csv -Path $soloDlpReportsFilePath -NoTypeInformation
        }

	    if (($soloEtrReports | Measure-Object).Count -gt 0)
	    {
            $soloEtrReportsFilePath = $reportFilePath + $Global:Strings.SoloETRReportsFilePath
            # -Headers $Global:Strings.HeadersSoloEtrReports
	        $soloEtrReports | Export-Csv -Path $soloEtrReportsFilePath -NoTypeInformation
	    }

        if (($restrictedEtrReports | Measure-Object).Count -gt 0)
        {
            $restrictedFilePath = $reportFilePath + $Global:Strings.RestrictedETRFilePath
            # -Headers $Global:Strings.HeadersSoloEtrReports
            $restrictedEtrReports | Export-Csv -Path $restrictedFilePath -NoTypeInformation
        }

        [Utils]::TryConvertCsvToExcel($reportFilePath)

        Write-Host "Generated Report present in file path: $reportFilePath"
    }

    static [void] TryConvertCsvToExcel([string]$ReportFolderName)
    {
        try
        {
            #Create Excel Com Object
            $excel = new-object -com excel.application
    
            #Disable alerts
            $excel.DisplayAlerts = $False
    
            #Show Excel application
            $excel.Visible = $False

            # All input files
            $inputFiles = Get-ChildItem -Path $ReportFolderName | Where-Object -Property "Name" -Match ".csv" | Select-Object -ExpandProperty Name
            
            $csvOrder = $Global:Strings.RestrictedETRFilePath, $Global:Strings.SoloETRReportsFilePath, $Global:Strings.SoloDLPReportsFilePath, $Global:Strings.ETRDLPParityReportsFilePath, $Global:Strings.AllETRReportsFilePath, $Global:Strings.AllDLPReportsFilePath, $Global:Strings.MainReportFilePath 
            $finalInputFiles = $csvOrder | Where-Object {$_ -in $inputFiles}

            #Find the number of CSVs being imported
            $count = ($inputFiles.count -1)
    
            #Add workbook
            $workbook = $excel.workbooks.Add()
    
            try 
            {
                #Remove other worksheets
                $workbook.worksheets.Item(2).delete()
                #After the first worksheet is removed,the next one takes its place
                $workbook.worksheets.Item(2).delete()
            }
            catch{}
    
            #Define initial worksheet number
            $i = 1

            $newCsvPath = Join-Path $ReportFolderName "CSVReports"
            if (-not (Test-Path $newCsvPath))
            {
                New-Item -Path $newCsvPath -ItemType directory
            }
    
            ForEach ($inputCsv in $finalInputFiles) {

                # Append the folder
                $input = Join-Path $ReportFolderName $inputCsv | Resolve-Path

                #If more than one file, create another worksheet for each file
                If ($i -gt 1) {
                    $workbook.worksheets.Add() | Out-Null
                }

                #Use the first worksheet in the workbook (also the newest created worksheet is always 1)
                $worksheet = $workbook.worksheets.Item(1)
            
                #Add name of CSV as worksheet name
                $worksheet.name = “$((GCI $input).basename)”
    
                #Open the CSV file in Excel, must be converted into complete path if no already done
                $tempcsv = $excel.Workbooks.Open($input)

                $tempsheet = $tempcsv.Worksheets.Item(1)
                #Copy contents of the CSV file
                $tempSheet.UsedRange.Copy() | Out-Null
                #Paste contents of CSV into existing workbook
                $worksheet.Paste()
    
                #Close temp workbook
                $tempcsv.close()
    
                #Select all used cells
                $range = $worksheet.UsedRange
    
                #Autofit the columns
                $range.EntireColumn.Autofit() | out-null
                $i++

                #Move CSV to new file
                Move-Item $input $newCsvPath
            }

            #Save spreadsheet 
            [string]$finalReportPath = Join-Path (Resolve-Path $ReportFolderName) "FinalReport.xlsx"
            $workbook.saveas($finalReportPath)
    
            Write-Log -WriteHost -Message "Report converted to XLSX for easy viewing!" -Fore
    
            #Close Excel
            $excel.quit()
        }
        catch
        {
            Write-Log -ErrorRecord $_ -Message "Could not convert CSV to XLSX" -Severity Error -WriteHost
            return;
        }
    }
}