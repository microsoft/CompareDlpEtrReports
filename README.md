# Compare-DlpEtrReports Powershell Module

This module helps compare parity between the DLP and ETR solutions offered by Microsoft for Exchange, by generating a report in the form of a spreadsheet, with different values

### How to install module

*Open powershell with admin elevation*
```powershell
Install-Module -Name PowerShellGet -Force -AllowClobber
```
*Close powershell and reopen with admin rights again*
```
Install-Module CompareDlpEtrReports -RequiredVersion 1.0.2
```

### How to run the cmdlet

The ReportName parameter is optional, but the StartDate, EndDate and AdminEmailAddress parameters are necessary.

```powershell
Compare-DlpEtrReports -StartDate [DateTime] -EndDate [DateTime] -AdminEmailAddress [EmailAddress] [-ReportName [string]]
```

Examples :
```powershell
Compare-DlpEtrReports -StartDate 04/31/2021 -EndDate 05/01/2021 -AdminEmailAddress admin@tenant.com

Compare-DlpEtrReports -ReportName "myReport.xlsx" -StartDate "04/29/2021 15:00:00" -EndDate "04/30/2021 11:00:00" -AdminEmailAddress admin@tenant.com
```

### How does this module work?

- Step 1: Download all the transport rules configured for the tenant
- Step 2: For each transport rule, find the corresponding policy / rule hits using the reporting cmdlets readily available for both DLP and ETR.
- Step 3: Aggregate the several reports generated for each rule.
- Step 4: Analyze each report, determine the actions taken and export a spreadsheet with data you can use to understand how DLP fairs with the legacy ETR solution.

*Note: Lastly, happy migration!*

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
