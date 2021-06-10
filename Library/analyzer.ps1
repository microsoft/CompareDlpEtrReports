Class Analyzer
{
    static [string] GetReasonForDlpMatchOnly([CommonReport]$dlpReport, $Rules)
    {
        [Object]$rule = $Rules | Where-Object -Property "Name" -eq $dlpReport.rule

        if ($rule.SetAuditSeverity -eq [Constants.Strings]::DoNotAuditString)
        {
            return "Auditing for ETR report is disabled"
        }

        return "Please open a support ticket"
    }

    static [string] GetReasonForEtrMatchOnly([CommonReport]$etrReport, [string[]]$allDlpRuleNames)
    {
        if ($null -ne $etrReport.tempReason)
        {
            return $etrReport.tempReason
        }

        if ($allDlpRuleNames -notcontains $etrReport.rule)
        {
            return "This rule hasn't migrated to DLP."
        }

        return "Please open a support ticket"
    }

    static [void] AnalyzeDlpReports($commonDlpReports, $commonEtrReports, $allDlpReports, $etrDlpMatchedReports, $soloDlpReports, $Rules)
    {
        [CommonReport]$commonDlpReport = $null

        # Evaluation of results DLP reports
        foreach ($ruleName in $commonDlpReports.Keys)
        {
            foreach ($commonDlpReport in $commonDlpReports[$ruleName])
            {
                # Add to all reports section
                $allDlpReports.Add($commonDlpReport)
                
                # Find ETR reports for the same rule and same message id (if it exists)
                $potentialEtrReports = $commonEtrReports[$ruleName] | Where-Object -Property "messageId" -eq $commonDlpReport.messageId

                $potentialEtrReportsCount = ($potentialEtrReports | Measure-Object).Count
                if (($null -eq $potentialEtrReports) -or ($potentialEtrReportsCount -eq 0))
                {
                    [string]$reasonForDlpMatchOnly = [Analyzer]::GetReasonForDlpMatchOnly($commonDlpReport, $Rules)
                    $soloDlpReports.Add([SoloReport]::new($commonDlpReport, $reasonForDlpMatchOnly))
                    continue;
                }

                if ($potentialEtrReportsCount -gt 1)
                {
                    $ruleName = $commonDlpReport.rule
                    $messageId = $commonDlpReport.messageId
                    Write-Log -Message "Logic gone wrong in grouping similar matches. Rule: $ruleName. MessageId: $messageId." -Severity Error
                    throw Exception::new("Bad logic")
                }
                else
                {
                    [CommonReport]$commonEtrReport = $potentialEtrReports
                }

                # Verify ETR and DLP report if it refers to the same message and same rule 
                if (-not([Analyzer]::VerifyEtrDlpReport($commonDlpReport, $commonEtrReport)))
                {   
                    $messageId = $commonDlpReport.messageId
                    $ruleName = $commonEtrReport.rule
                    Write-Log -Message "DLP and ETR Report don't match for messageId: $messageId, rule: $ruleName" -Severity Warning
                }

                # The actions that each agent took
                $etrActions = $commonEtrReport.actions
                $dlpActions = $commonDlpReport.actions

                # Get the combined ETR-DLP report
                $combinedReport = [CombinedReport]::new($commonDlpReport, $etrActions, $dlpActions)

                # Add to the matched report
                $etrDlpMatchedReports.Add($combinedReport)
            }
        }
    }

    static [void] AnalyzeEtrReports($commonDlpReports, $commonEtrReports, $allEtrReports, $soloEtrReports, $restrictedEtrReports, $allDlpRuleNames)
    {
        [CommonReport]$commonEtrReport = $null

        # ETR Report those are not there.
        foreach ($ruleName in $commonEtrReports.Keys)
        {
            foreach ($commonEtrReport in $commonEtrReports[$ruleName])
            {
                # Add each ETR report to the list of all ETR reports
                $allEtrReports.Add($commonEtrReport)

                # Find the DLP report which matches the ETR report for same rule and message Id
                $commonDlpReport = $commonDlpReports[$commonEtrReport.rule] | Where-Object -Property "messageId" -eq $commonEtrReport.messageId

                if (($null -eq $commonDlpReport) -or ($commonDlpReport | Measure-Object).Count -eq 0)
                {
                    [string]$reasonForEtrOnly = [Analyzer]::GetReasonForEtrMatchOnly($commonEtrReport, $allDlpRuleNames)

                    $soloReport = [SoloReport]::new($commonEtrReport, $reasonForEtrOnly)
                    if (-not $reasonForEtrOnly.Contains("Restrictive"))
                    {
                        $soloEtrReports.Add($soloReport)
                    }
                    else
                    {
                        $restrictedEtrReports.Add($soloReport)
                    }
                }
            }
        }
    }
    
    static [bool]VerifyEtrDlpReport([CommonReport]$DlpReport, [CommonReport]$EtrReport)
    {
        Write-Log -Message "Verifying DLP and ETR Report for " -Severity Debug
        if ($DlpReport.messageId -eq $EtrReport.messageId)
        {
            if ($DlpReport.sender -eq $EtrReport.sender)
            {
                if ($DlpReport.siType -eq $EtrReport.siType)
                {
                    if ($DlpReport.siCount -eq $EtrReport.siCount)
                    {
                        if ($DlpReport.rule -eq $EtrReport.rule)
                        {
                            if ($DlpReport.policy -eq $EtrReport.policy)
                            {
                                return $true;
                            }
                            else {
                                $DlpValue = $DlpReport.policy
                                $EtrValue = $EtrReport.policy
                                Write-Log "Policy $DlpValue and $EtrValue doesn't match"
                            }
                        }
                        else {
                            $DlpValue = $DlpReport.rule
                            $EtrValue = $EtrReport.rule
                            Write-Log "Rule $DlpValue and $EtrValue doesn't match"
                        }
                    }
                    else {
                        $DlpValue = $DlpReport.sicount
                        $EtrValue = $EtrReport.sicount
                        Write-Log "Sensitive information count $DlpValue and $EtrValue doesn't match"
                    }
                }
                else {
                    $DlpValue = $DlpReport.siType
                    $EtrValue = $EtrReport.siType
                    Write-Log "Sensitive type $DlpValue and $EtrValue doesn't match"
                }
            }
            else {
                $DlpValue = $DlpReport.sender
                $EtrValue = $EtrReport.sender
                Write-Log "Senders $DlpValue and $EtrValue doesn't match"
            }
        }
        else {
            $DlpValue = $DlpReport.messageId
            $EtrValue = $EtrReport.messageId
            Write-Log "MessageID $DlpValue and $EtrValue doesn't match"
        }
        
        return $false;
    }
}