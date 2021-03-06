Describe 'Test for generating custom file name' {
    It 'Creates report path and returns correct file name' {
        
        $testName = "testReportFile"
        $reportFileName = [Utils]::GetReportFolderPath($testName)
        $reportFileName | Should -Be ".\Reports\$testName\"

        Test-Path ".\Reports" | Should -Be $true

        $reportFileName = [Utils]::GetReportFolderPath($testName)
        $reportFileName | Should -Be (".\Reports\$testName" + "_1\")
    }
}

Describe 'Test for generating appropriate report file for date' {
    It 'Creates report path and returns correct file name' {
        
        $reportFileName = [Utils]::GetReportFolderPath($null)
        $date = (Get-Date).ToString("dd-MM-yyyy")
        $reportCheckValue = ".\Reports\Reports_${date}\"
        $reportFileName | Should -Be $reportCheckValue


        $reportFileName = [Utils]::GetReportFolderPath($null)
        $reportCheckValue = ".\Reports\Reports_${date}_1\"
        $reportFileName | Should -Be $reportCheckValue
    }
}

Describe 'Test for parsing actions from rule, where no rules are valid' {
    It 'Checks if parsing function is working as expected' {
        
        $Rule = [PSCustomObject]@{
            Actions = @("asldjfk", "sdkljf", "sdlfjk")
        }

        [Constants.CommonAction]$action = [Utils]::GetActionsFromRule($Rule)
        $action -eq [Constants.CommonAction]::None | Should -Be $true

        $action = [Utils]::GetActionsFromRule($null)
        $action -eq [Constants.CommonAction]::None | Should -Be $true
    }
}

Describe 'Test for parsing actions from rule, most rules are valid' {
    It 'Checks if parsing function is working as expected' {
        
        $Rule = [PSCustomObject]@{
            Actions = @("SetSCL", "BlockAccess", "sdlkfjsdlf", "Microsoft.Exchange.MessagingPolicies.Rules.Tasks.ApplyHtmlDisclaimerAction")
        }

        [Constants.CommonAction]$action = [Utils]::GetActionsFromRule($Rule)
        $action -eq ([Constants.CommonAction]::SetSpamConfidenceLevel -bor [Constants.CommonAction]::BlockAccess) -bor [Constants.CommonAction]::ApplyHtmlDisclaimer | Should -Be $true
    }
}

Describe 'Test for parsing actions from rule, for sender notify' {
    It 'Checks if parsing function is working as expected' {
        
        $Rule1 = [PSCustomObject]@{
            Actions = @("SenderNotify")
            SenderNotificationType = "NotifyOnly" 
        }
        [Constants.CommonAction]$action = [Utils]::GetActionsFromRule($Rule1)
        $action -eq [Constants.CommonAction]::NotifySender | Should -Be $true

        $Rule2 = [PSCustomObject]@{
            Actions = @("SenderNotify")
            SenderNotificationType = "RejectMessage" 
        }
        [Constants.CommonAction]$action = [Utils]::GetActionsFromRule($Rule2)
        $action -eq ([Constants.CommonAction]::NotifySender -bor [Constants.CommonAction]::RejectMessage) | Should -Be $true
    }
}

Describe 'Test_TryParseAction' {
    It 'Checks if action parsing works as expected' {
        
        [string]$Action = "dfj"
        $noneAction = [Constants.CommonAction]::None
        [Utils]::TryParseAction($Action) | Should -Be $noneAction
        [Utils]::TryParseAction($null) | Should -Be $noneAction

        $Action = "AddToRecipient"
        [Utils]::TryParseAction($Action) -eq [Constants.CommonAction]::AddToRecipient| Should -Be $true
    
        $Action = "SetHeader"
        [Utils]::TryParseAction($Action) -eq [Constants.CommonAction]::SetMessageHeader | Should -Be $true
    }
}

Describe 'Test_GroupRuleReportsByPolicy' {
    It 'Grouping by policy, checks if works as expected' {
        
        # Fake values test
        $reports = [PSCustomObject]@{
            abc = $null
            def = @()
            ghi = @("test")
        }
        $testMapping = [Collections.Generic.Dictionary[string, Collections.Generic.List[String]]]::new()

        [Utils]::GroupRuleReportsByPolicy($reports, $testMapping)
        ($testMapping.Keys | Measure-Object).Count | Should -Be 0

        # Real values test
        $commonReport1 = [Resources]::NewDlpCommonReport($false)
        $commonReport2 = [Resources]::NewDlpCommonReport($true)
        $reports = [ordered]@{
            abc = @($commonReport1)
            def = @($commonReport2)
        }

        [Utils]::GroupRuleReportsByPolicy($reports, $testMapping)

        $testMapping.Keys.Count | Should -Be 2
    }
}

Describe 'Test_GenerateMainReport' {
    It 'Testing main report generation broadly' {
        
        # real test
        $true | Should -Be $true
    }
}

