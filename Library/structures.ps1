Class CommonReport
{
    [string]$messageId
    [string]$subject
    [string]$sender
    [string]$receivers
    [string]$policy
    [string]$rule
    [string]$siType
    [int]$siCount
    [datetime]$date
    [Constants.CommonAction]$actions
    [string]$tempReason
    [string]$tempReasonAction

    # Constructor
    CommonReport([Object]$specificReport, [ReportType]$reportType, [Constants.CommonAction]$action)
    {
        if ($reportType -eq [ReportType]::DlpReport)
        {
            $this.subject = $specificReport.Title;
            $this.messageId = $specificReport.Location;
            $this.sender = $specificReport.Actor;
            $this.receivers = $specificReport.Recipients
            $this.policy = $specificReport.DlpCompliancePolicy
            $this.rule = $specificReport.DlpComplianceRule
            $this.siType = $specificReport.SensitiveInformationType
            $this.siCount = $specificReport.SensitiveInformationCount
            $this.date = $specificReport.LastModifiedTime
        }  
        else
        {
            $this.subject = $specificReport.Subject;
            $this.messageId = $specificReport.MessageID;
            $this.sender = $specificReport.SenderAddress;
            $this.receivers = $specificReport.RecipientAddress
            if ($null -eq $specificReport.DLPPolicy)
            {
                $this.policy = $Global:IndependentRulePolicy
            }
            else
            {
                $this.policy = $specificReport.DLPPolicy
            }
            
            $this.rule = $specificReport.TransportRule
            $this.siType = $specificReport.SensitiveInformationType
            $this.siCount = $specificReport.SensitiveInformationCount
            $this.date = $specificReport.Date

            $AllBlockedMessages = $Global:AllMessagesActedUponByBlockingRules

            if (($null -ne $AllBlockedMessages) -and $AllBlockedMessages.ContainsKey($this.messageId))
            {
                [BlockedRuleDetails]$blockedRuleDetails = $Global:AllMessagesActedUponByBlockingRules[$this.messageId]  
                if (-not($blockedRuleDetails.IsOverridable -and ($specificReport.UserAction -eq "or" -or $specificReport.UserAction -eq "fp")))
                {
                    $ruleName = $blockedRuleDetails.RuleName
                    $this.tempReason = "Restrictive actions taken by rule '$ruleName' prevented DLP from acting"
                }
            }
        }

        $this.actions = $action
    }

    # Copy constructor
    CommonReport([CommonReport]$commonReport)
    {
        $this.messageId = $commonReport.messageId
        $this.subject = $commonReport.subject
        $this.sender = $commonReport.sender
        $this.receivers = $commonReport.receivers
        $this.policy = $commonReport.policy
        $this.rule = $commonReport.rule
        $this.siType = $commonReport.siType
        $this.siCount = $commonReport.siCount
        $this.date = $commonReport.date
        $this.actions = $commonReport.actions
    }
}

Class CombinedReport: CommonReport
{
    [Constants.CommonAction]$etrActions
    [Constants.CommonAction]$dlpActions

    CombinedReport([CommonReport]$commonReport, [Constants.CommonAction]$etrActions, [Constants.CommonAction]$dlpActions): base($commonReport)
    {
        $this.etrActions = $etrActions
        $this.dlpActions = $dlpActions
    }
}

Class SoloReport: CommonReport
{
    [string] $reason
    SoloReport([CommonReport]$commonReport, [string]$reason): base($commonReport)
    {
        $this.reason = $reason
    }
}

Class MainReport
{
    [string]$policyName
    [string]$ruleName
    [string]$etrMatchesWithRestriction
    [string]$etrMatches
    [string]$dlpMatches
    [string]$percentageMatch

    MainReport([string]$ruleName, [int]$dlpMatches, [int]$etrMatchesWithoutRestriction, [int]$etrMatchesWithRestriction)
    {
        $this.policyName = ""
        $this.dlpMatches = $dlpMatches.ToString()
        $this.etrMatches = $etrMatchesWithoutRestriction.ToString()
        $this.ruleName = $ruleName
        $this.CalculatePercentageMatch($dlpMatches, $etrMatchesWithoutRestriction)
        $this.etrMatchesWithRestriction = $etrMatchesWithRestriction.ToString()
    }

    MainReport([string]$policyName)
    {
        $this.policyName = $policyName
    }

    MainReport([string]$policyName, [string]$ruleName)
    {
        $this.policyName = $policyName
        $this.ruleName = $ruleName
    }

    [void] CalculatePercentageMatch([int]$dlpMatches, [int]$etrMatches)
    {
        if ($etrMatches -eq 0)
        {
            $this.percentageMatch = "-"
        }
        else
        {
            $this.percentageMatch = ([math]::Round(($dlpMatches / $etrMatches)*100, 2)).ToString()
        }
    }
}

Class BlockedRuleDetails
{
    [string]$RuleName
    [string]$ActionsTaken
    [bool]$IsOverridable

    BlockedRuleDetails([string]$RuleName, [string]$ActionsTaken, [bool]$IsOverridable)
    {
        $this.RuleName = $RuleName
        $this.ActionsTaken = $ActionsTaken
        $this.IsOverridable = $IsOverridable
    }
}
