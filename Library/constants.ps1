$Global:IndependentRulePolicy = "Independent Rule (not part of a DLP policy)"

[Flags()]enum ReportType{
    DlpReport
    EtrReport
}

$Actions = @"
    using System;
    namespace Constants{

        [Flags]
        public enum CommonAction : long
        {
            None                            = 0x0000000000000000L,
            AddBccRecipient                 = 0x0000000000000001L,
            AddCcRecipient                  = 0x0000000000000002L,
            AddManagerAsRecipient           = 0x0000000000000004L,
            AddToRecipient                  = 0x0000000000000008L,
            ApplyClassification             = 0x0000000000000010L,
            ApplyHtmlDisclaimer             = 0x0000000000000020L,
            DeleteMessage                   = 0x0000000000000040L,
            GenerateIncidentReport          = 0x0000000000000100L,
            ModerateMessageByManager        = 0x0000000000000200L,
            ModerateMessageByUser           = 0x0000000000000400L,
            NotifySender                    = 0x0000000000001000L,
            PrependSubject                  = 0x0000000000002000L,
            Quarantine                      = 0x0000000000004000L,
            RedirectMessage                 = 0x0000000000008000L,
            RejectMessage                   = 0x0000000000010000L,
            RemoveMessageHeader             = 0x0000000000020000L,
            RequireTLS                      = 0x0000000000040000L,
            RightsProtectMessage            = 0x0000000000080000L,
            RouteMessageUsingConnector      = 0x0000000000100000L,
            SetAuditSeverityHigh            = 0x0000000000200000L,
            SetAuditSeverityLow             = 0x0000000000400000L,
            SetAuditSeverityMedium          = 0x0000000000800000L,
            SetMessageHeader                = 0x0000000001000000L,
            SetSpamConfidenceLevel          = 0x0000000002000000L,
            StopRuleProcessing              = 0x0000000004000000L, 
            BlockAccess                     = 0x0000000008000000L,
            NotifyUser                      = 0x0000000010000000L, 
            GenerateNotification            = 0x0000000020000000L, 
            BypassMessage                   = 0x0000000040000000L, 
            BypassSystemMessage             = 0x0000000080000000L, 
            ContentReplaced                 = 0x0000000100000000L,
            InfectedAllowed                 = 0x0000000200000000L,
            AllowRedirect                   = 0x0000000400000000L,
            BlockRedirect                   = 0x0000000800000000L,
            ReplaceRedirect                 = 0x0000001000000000L,
            SubmissionBulkSetting           = 0x0000002000000000L,
            SubmissionETR                   = 0x0000004000000000L,
            SubmissionASF                   = 0x0000008000000000L,
            SubmissionBlockAllow            = 0x0000010000000000L,
            Allow                           = 0x0000020000000000L,
            DynamicReplaced                 = 0x0000040000000000L,
            CaughtAsSpam                    = 0x0000080000000000L,
            GoodMail                        = 0x0000100000000000L,
            TotalMsg                        = 0x0000200000000000L,
            ReleaseMsg                      = 0x0000400000000000L,
            ReportMsg                       = 0x0000800000000000L,
            ReleaseReportMsg                = 0x0001000000000000L,

            // Custom
            RemoveOME                       = 0x0002000000000000L,
        }

        public class Strings{
            public static string DoNotAuditString = "DoNotAudit";
        }
    }
"@

Add-Type -TypeDefinition $Actions -Language CSharp	

$Global:EtrActionToCommonActionMap = [ordered]@{
    
    AddToRecipient                 = [Constants.CommonAction]::AddToRecipient ;
    CopyTo                         = [Constants.CommonAction]::AddCcRecipient ;
    BlindCopyTo                    = [Constants.CommonAction]::AddBccRecipient ;
    AddManagerAsRecipientType      = [Constants.CommonAction]::AddManagerAsRecipient; 
    ApplyHtmlDisclaimer            = [Constants.CommonAction]::ApplyHtmlDisclaimer ; 
    DeleteMessage                  = [Constants.CommonAction]::BlockAccess
    GenerateIncidentReport         = [Constants.CommonAction]::GenerateIncidentReport ;
    Halt                           = [Constants.CommonAction]::StopRuleProcessing; 
    ModerateMessageByManager       = [Constants.CommonAction]::ModerateMessageByManager ;
    ModerateMessageByUser          = [Constants.CommonAction]::ModerateMessageByUser ; 
    PrependSubject                 = [Constants.CommonAction]::PrependSubject ;
    Quarantine                     = [Constants.CommonAction]::ModerateMessageByManager ;
    RedirectMessage                = [Constants.CommonAction]::RedirectMessage ;
    RejectMessage                  = [Constants.CommonAction]::RejectMessage ;
    RemoveHeader                   = [Constants.CommonAction]::RemoveMessageHeader  ;
    RightsProtectMessage           = [Constants.CommonAction]::RightsProtectMessage ;
    SenderNotify                   = [Constants.CommonAction]::NotifySender ;  # Notify User / Notify Sender ?
    SetHeader                      = [Constants.CommonAction]::SetMessageHeader ;
    SetSCL                         = [Constants.CommonAction]::SetSpamConfidenceLevel ;
    ReportSeverityLevelHigh        = [Constants.CommonAction]::SetAuditSeverityHigh ;
    ReportSeverityLevelLow         = [Constants.CommonAction]::SetAuditSeverityLow ;
    ReportSeverityLevelMed         = [Constants.CommonAction]::SetAuditSeverityMed ;

    #Custom ones
    RemoveOMEv2                    = [Constants.CommonAction]::RemoveOME ;
    Decrypt                        = [Constants.CommonAction]::RemoveOME ;
    RemoveRMSAttachmentEncryption  = [Constants.CommonAction]::RemoveOME ;
    DecryptMessage                 = [Constants.CommonAction]::RemoveOME ;
    EncryptMessage                 = [Constants.CommonAction]::RightsProtectMessage ;
    RightsProtectMessageCustomization          = [Constants.CommonAction]::RightsProtectMessage ;

    # Only ETR rule
    SetAuditSeverity               = [Constants.CommonAction]::None ;

    # Do not migrate
    ApplyClassification            = [Constants.CommonAction]::ApplyClassification ;
    RouteMessageOutboundConnector  = [Constants.CommonAction]::RouteMessageUsingConnector ;
    RouteMessageOutboundRequireTls = [Constants.CommonAction]::RequireTLS ;
}

$Global:EtrBlockingActions = 
    [Constants.CommonAction]::BlockAccess `
     -bor [Constants.CommonAction]::DeleteMessage `
     -bor [Constants.CommonAction]::ModerateMessageByManager `
     -bor [Constants.CommonAction]::ModerateMessageByUser `
     -bor [Constants.CommonAction]::RejectMessage   

$Global:EtrAuditingActions = 
    [Constants.CommonAction]::SetAuditSeverityLow `
    -bor [Constants.CommonAction]::SetAuditSeverityMedium `
    -bor [Constants.CommonAction]::SetAuditSeverityHigh