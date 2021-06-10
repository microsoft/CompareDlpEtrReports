Describe 'GetNonBlockingTransportRules_Test' {
    It ' Returns only the rules with non-blocking actions' {
        
        $transportRules = Import-Clixml ".\Resources\TransportRules.xml"
        $nonBlockingRules = [CmdletUtils]::GetNonBlockingTransportRules($transportRules)

        # Check that all the rules have only non-blocking actions
        foreach ($NonblockingRule in $nonBlockingRules)
        {
            $actions = [Utils]::GetActionsFromRule($NonblockingRule)
            $actions = $actions -band $Global:EtrBlockingAcions

            if ($actions -ne [Constants.CommonAction]::None)
            {
                $false | Should -Be $true
            }
        }

        # null test
        [CmdletUtils]::GetNonBlockingTransportRules($null) | Measure-Object | Select-Object -ExpandProperty "Count" | Should -Be 0
    }
}

Describe 'GroupTransportRulesByPolicy_Test' {
    It ' Groups transport rules by policy' {
        
        $transportRules = Import-Clixml ".\Resources\TransportRules.xml"
        $nonBlockingRules = [CmdletUtils]::GetNonBlockingTransportRules($transportRules)
        $NonBlockingRuleNames = $nonBlockingRules | Select-Object -ExpandProperty "Name"

        # Rules with blocking actions
        $BlockingRules = $transportRules | Where-Object -Property Name -NotIn $NonBlockingRuleNames
        [Collections.Generic.Dictionary[String, Collections.Generic.List[String]]]$perPolicyRules = [CmdletUtils]::GroupTransportRulesByPolicy($BlockingRules)
        $perPolicyRules.Keys.Count | Should -Be 1

        # null check
        $perPolicyRules = [CmdletUtils]::GroupTransportRulesByPolicy($null)
        $perPolicyRules.Keys.Count | Should -Be 0

        $newTransportRule = Import-Clixml ".\Resources\TransportRule.xml"
        $newTransportRule.DlpPolicy = "newPolicy"
        $BlockingRules += $newTransportRule
        $perPolicyRules = [CmdletUtils]::GroupTransportRulesByPolicy($BlockingRules)
        $perPolicyRules.Keys.Count | Should -Be 2

    }
}

Describe 'GetBlockedMessages_Test' {
    It ' Gets all the blocked messages by some rule' {
        
        $transportRules = Import-Clixml ".\Resources\TransportRules.xml"
        $nonBlockingRules = [CmdletUtils]::GetNonBlockingTransportRules($transportRules)
        $NonBlockingRuleNames = $nonBlockingRules | Select-Object -ExpandProperty "Name"

        # Rules with blocking actions
        $BlockingRules = $transportRules | Where-Object -Property Name -NotIn $NonBlockingRuleNames
        [Collections.Generic.Dictionary[String, Collections.Generic.List[String]]]$perPolicyRules = [CmdletUtils]::GroupTransportRulesByPolicy($BlockingRules)
        $perPolicyRules.Keys.Count | Should -Be 1

        # null check
        $perPolicyRules = [CmdletUtils]::GroupTransportRulesByPolicy($null)
        $perPolicyRules.Keys.Count | Should -Be 0

        $newTransportRule = Import-Clixml ".\Resources\TransportRule.xml"
        $newTransportRule.DlpPolicy = "newPolicy"
        $BlockingRules += $newTransportRule
        $perPolicyRules = [CmdletUtils]::GroupTransportRulesByPolicy($BlockingRules)
        $perPolicyRules.Keys.Count | Should -Be 2

    }
}

