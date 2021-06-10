Describe 'Test_GenerateCommonReport' {
    It 'Only filters a particular rule and generates common reports' {

        # Null case
        [Collections.Generic.List[CommonReport]]$commonReports = [Aggregator]::GenerateCommonReports($null, [ReportType]::DlpReport, $null)
        $commonReports.Count | Should -Be 0

        $DlpReports = Import-Clixml ".\Resources\DlpReportsSameMessage.xml"
        $commonReports = [Aggregator]::GenerateCommonReports($DlpReports, [ReportType]::DlpReport, "Location")
        $commonReports.Count | Should -Be 1

        $DlpReports = Import-Clixml ".\Resources\DlpReportsDifferentMessages.xml"
        $commonReports = [Aggregator]::GenerateCommonReports($DlpReports, [ReportType]::DlpReport, "Location")
        $commonReports.Count | Should -Be 3

        $DlpReports[3] = $null
        $commonReports = [Aggregator]::GenerateCommonReports($DlpReports, [ReportType]::DlpReport, "Location")
        $commonReports.Count | Should -Be 3

    }
}

Describe 'Test_FiterAndGroupReports' {
    It 'Generating common report remains the most important task of this module' {

        # Null case
        [Collections.Generic.List[CommonReport]]$commonReports = [Aggregator]::GenerateCommonReports($null, [ReportType]::DlpReport, $null)
        $commonReports.Count | Should -Be 0

        $DlpReports = Import-Clixml ".\Resources\DlpReportsSameMessage.xml"
        $commonReports = [Aggregator]::FilterAndGroupReportsBasedOnRule($DlpReports, "AppendCheck", [ReportType]::DlpReport, "Location")
        $commonReports.Count | Should -Be 1

        $DlpReports = Import-Clixml ".\Resources\DlpReportsDifferentMessages.xml"
        $commonReports = [Aggregator]::FilterAndGroupReportsBasedOnRule($DlpReports, "AppendCheck", [ReportType]::DlpReport, "Location")
        $commonReports.Count | Should -Be 1

        $commonReports = [Aggregator]::FilterAndGroupReportsBasedOnRule($DlpReports, "RandomRule", [ReportType]::DlpReport, "Location")
        $commonReports.Count | Should -Be 0
    }
}