Class Resources
{
    static [Object] $TransportRule
    static [Object] $DlpReport
    static [Object] $EtrReport

    static Resources()
    {
        # Directly from the cmdlet
        [Resources]::TransportRule = Import-Clixml .\Resources\TransportRule.xml
        [Resources]::EtrReport = Import-Clixml .\Resources\EtrReport.xml
        [Resources]::DlpReport = Import-Clixml .\Resources\DlpReport.xml
    }

    static [CommonReport]NewEtrCommonReport([Boolean]$Randomize)
    {
        [CommonReport]$commonReport = [CommonReport]::new([Resources]::EtrReport, [ReportType]::EtrReport, [Constants.CommonAction]::AddToRecipient)
        return [Resources]::Randomize($commonReport, $Randomize)
    }

    static [CommonReport]NewDlpCommonReport([Boolean]$Randomize)
    {
        [CommonReport]$commonReport = [CommonReport]::new([Resources]::DlpReport, [ReportType]::DlpReport, [Constants.CommonAction]::AddToRecipient)
        return [Resources]::Randomize($commonReport, $Randomize)
    }

    static [CommonReport]Randomize([CommonReport]$commonReport, [Boolean]$Randomize)
    {
        if (-not $Randomize)
        {
            return $commonReport
        }

        $commonReport.rule = -join ((65..90) + (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_})
        $commonReport.policy = -join ((65..90) + (97..122) | Get-Random -Count 10 | ForEach-Object {[char]$_})

        return $commonReport
    }
}

