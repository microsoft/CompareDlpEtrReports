Class CmdletUtils
{
    static [void] GetOAuthTokenForCmdlets([string]$adminUserName)
    {    
        Write-Log -Message "Attempting to create IP session for $adminUserName"

        if (-not (Get-Command 'Get-TransportRule' -errorAction SilentlyContinue)) 
        {
            Connect-IPPSSession -UserPrincipalName $adminUserName
        }

	    Write-Log -Message "Attempting to create Exchange Online session"
	
        if (-not (Get-Command 'Get-DlpDetailReport' -errorAction SilentlyContinue)) 
        {
            Connect-ExchangeOnline -UserPrincipalName $adminUserName
        } 
        
        Write-Log -Message "Connected sessions, ready to call cmdlets" -Checkpoint
    }

    static [Object[]] GetTransportRules()
    {
        Write-Log -Message "Trying to retrieve transport rules using Get-TransportRule"
        
        [Object[]]$TransportRules = Get-TransportRule
        [int]$TransportRulesCount = ($TransportRules | Measure-Object).Count

        Write-Log -Message "Retrieved $TransportRulesCount transport rules using cmdlet" -Checkpoint

        return $TransportRules
    }

    static [String[]] GetDlpRulesNames()
    {
        Write-Log -Message "Trying to retrieve DLP rules using Get-DlpComplianceRule"
        
        [Object[]]$DlpRules = Get-DlpComplianceRule
        [String[]]$AllDlpRuleNames = $DlpRules | Where-Object -Property "Mode" -ne "PendingDeletion" | Select-Object -ExpandProperty "Name"
        [int]$DlpRulesCount = ($DlpRules | Measure-Object).Count

        Write-Log -Message "Retrieved $DlpRulesCount DLP rules using cmdlet" -Checkpoint

        return $AllDlpRuleNames
    }
    
    static [Object[]] GetNonBlockingTransportRules($Rules)
    {
        $NonBlockingRules = [Collections.Generic.List[Object]]::new()

        #Removing Rules with Restrictive Actions
        foreach ($Rule in $Rules)
        {
            $RuleName = $Rule.Name
            Write-Log -Message "Trying to parse actions for rule: $RuleName"
            [Constants.CommonAction]$reportActions = [Utils]::GetActionsFromRule($Rule)   
            Write-Log -Message "Successfully parsed actions for rule: $RuleName"

            if (($reportActions -band $Global:EtrBlockingActions) -eq [Constants.CommonAction]::None)
            {
                $NonBlockingRules.Add($Rule)
            }   
        }

        return $NonBlockingRules
    }

    static [Collections.Generic.Dictionary[String, Collections.Generic.List[String]]] GroupTransportRulesByPolicy($BlockingRules)
    {
        $BlockingRulesPerPolicy = New-Object System.Collections.Generic.Dictionary"[String,Collections.Generic.List[String]]"

        Write-Log -Message "Trying to group rules by policy (for blocking rules)" -Checkpoint
        foreach ($BlockingRule in $BlockingRules)
        {
            $ruleName = $BlockingRule.Name
            $policyName = $BlockingRule.DlpPolicy
            if ($null -eq $policyName)
            {
                $policyName = $Global:IndependentRulePolicy
            }

            if (-not $BlockingRulesPerPolicy.ContainsKey($policyName))
            {
                $newRuleList = [Collections.Generic.List[String]]::new()
                $BlockingRulesPerPolicy.Add($policyName, $newRuleList)
            }

            Write-Log -Message "Added rule: $ruleName under policy: $policyName"
            $BlockingRulesPerPolicy[$policyName].Add($ruleName)
        }

        return $BlockingRulesPerPolicy
    }

    static [Collections.Generic.Dictionary[String, BlockedRuleDetails]] GetMessagesActedByBlockingRules($BlockingRules, $EtrCommand)
    {
        Write-Log -Message "Trying to group rules by policy (for blocking rules)"

        $BlockedMessages = [Collections.Generic.Dictionary[String, BlockedRuleDetails]]::new()
        foreach ($BlockingRule in $BlockingRules)
        {
            $BlockingRuleName = $BlockingRule | Select-Object -ExpandProperty "Name"

            Write-Log -Message "Finding reports for blocked rule: $BlockingRuleName"

            [string] $command = $EtrCommand + '"' + $BlockingRuleName + '"'
            $CmdletReports = Invoke-Expression $command

            $MessagesActedOn = $CmdletReports | Where-Object -Property "TransportRule" -eq $BlockingRuleName | Select-Object -ExpandProperty "MessageId"
            Write-Log -Message "Successfully found messages acted on by $BlockingRuleName"

            [CmdletUtils]::AddMessagesToBlockedMessageDictionary($MessagesActedOn, $BlockedMessages, $BlockingRule, $BlockingRuleName)
        }

        Write-Log "Returning the blocking messages that were acted on by different rules" -Checkpoint
        return $BlockedMessages
    } 

    static [void] AddMessagesToBlockedMessageDictionary($MessagesActedOn, $BlockedMessages, $BlockingRule, $BlockingRuleName)
    {
        [Constants.CommonAction] $ruleActions = [Utils]::GetActionsFromRule($BlockingRule)

        foreach($MessageActedOn in $MessagesActedOn)
        {
            if (-not $BlockedMessages.ContainsKey($MessageActedOn))
            {
                $IsOverrideable = $true
                if (($ruleActions -band [Constants.CommonAction]::NotifySender) -ne [Constants.CommonAction]::None)
                {
                    $notificationType = $BlockingRule.SenderNotificationType
                    if (($notificationType -eq "RejectMessage") -or ($notificationType -eq "NotifyUser"))
                    {
                        $IsOverrideable = $false
                    }
                }

                $blockingActionsTaken = $ruleActions -band $Global:EtrBlockingActions
                $BlockedMessages.Add($MessageActedOn, [BlockedRuleDetails]::new($BlockingRuleName, $blockingActionsTaken, $IsOverrideable))
            }
        }
    }
}