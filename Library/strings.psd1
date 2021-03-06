@{
    MainReportFilePath = "MainReport.csv"
    HeadersMainReport = "Policy", "Rule", "All ETR Matches", "ETR Matches without restriction", "Dlp Matches", "Percentage Match"
    
    AllDLPReportsFilePath = "AllDLPReports.csv"
    HeadersAllDlpReports = "Message ID", "Subject", "Sender", "Receivers", "Policy", "Rule", "Sensitive Info Type", "SI count", "Date", "Actions taken"
    
    AllETRReportsFilePath = "AllETRReports.csv"
    HeadersAllEtrReports = "Message ID", "Subject", "Sender", "Receivers", "Policy", "Rule", "Sensitive Info Type", "SI count", "Date", "Actions taken"
    
    ETRDLPParityReportsFilePath = "ETRDLPParityReports.csv"
    HeadersMatchedReports = "ETR actions taken", "DLP actions taken", "Message ID", "Subject", "Sender", "Receivers", "Policy", "Rule", "Sensitive Info Type", "SI count", "Date"
    
    SoloDLPReportsFilePath = "SoloDLPReports.csv"
    HeadersSoloDlpReports = "Reason for ETR not matching", "Message ID", "Subject", "Sender", "Receivers", "Policy", "Rule", "Sensitive Info Type", "SI count", "Date", "Actions taken"
    
    SoloETRReportsFilePath = "SoloETRReports.csv"
    HeadersSoloEtrReports = "Reason for DLP not matching", "Message ID", "Subject", "Sender", "Receivers", "Policy", "Rule", "Sensitive Info Type", "SI count", "Date", "Actions taken"
    
    RestrictedETRFilePath = "RestrictedETRAgent.csv"
}