#################################################################################
#																				#
#								run_me.ps1										#
#																				#
#################################################################################
<#
.NOTES
The idea of this script is replace the MPARR_Collector.ps1 script as a task under
task scheduler in case that you need to execute the previous script with some attributes

MPARR Collector have the capability to be executed with some filters that can be set to 
"Contains" or "NotContains", this filters are set in the schemas.json file.

With his filters the MPARR_Collector can be used to download all the Operations (activities)
that contains certain value or not download if contains another value

Finally MPARR_Collector can be executed like this:
PS C:\MPARR Collector> .\MPARR_Collector.ps1 -FilterAuditExchange "DLP|Sensitivity" -FilterAuditSharePoint "DLP|Sensitivity" -FilterAuditGeneral "DLP|Sensitivity" -FilterDLPAll "DLP|Sensitivity"

In this example the MPARR_Collector will request only operations that contains DLP or Sensitivity, that can match with some operations like this ones:
CreateDlpPolicy
DLPRuleMatch
Get-DlpCompliancePolicy
Get-DlpComplianceRule
Get-DlpDetailReport
Get-DlpDetectionsReport
Get-DlpEdmSchema
Get-DlpKeywordDictionary
Get-DlpSensitiveInformationType
Get-DlpSensitiveInformationTypeRulePackage
New-DlpCompliancePolicy
New-DlpComplianceRule
New-DlpEdmSchema
New-DlpSensitiveInformationTypeRulePackage
Remove-DlpComplianceRule
Set-DlpCompliancePolicy
Set-DlpComplianceRule
Set-DlpEdmSchema
Set-DlpSensitiveInformationTypeRulePackage
UpdateDlpPolicy

OR 

AutoSensitivityLabelRuleMatch
Get-AutoSensitivityLabelPolicy
Get-AutoSensitivityLabelRule
New-AutoSensitivityLabelPolicy
New-AutoSensitivityLabelRule
Remove-AutoSensitivityLabelPolicy
SensitivityLabelApplied
SensitivityLabelChanged
SensitivityLabeledFileOpened
SensitivityLabeledFileRenamed
SensitivityLabelPolicyMatched
SensitivityLabelRemoved
SensitivityLabelUpdated
Set-AutoSensitivityLabelPolicy
Set-AutoSensitivityLabelRule

.NOTES
Please use your right path
#>
cd 'C:\MPARR Collector\'
.\MPARR_Collector.ps1 -FilterAuditExchange "DLP|Sensitivity" -FilterAuditSharePoint "DLP|Sensitivity" -FilterAuditGeneral "DLP|Sensitivity" -FilterDLPAll "DLP|Sensitivity"