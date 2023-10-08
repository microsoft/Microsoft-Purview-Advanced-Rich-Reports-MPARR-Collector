# Kusto Query for MIP Scanner

> [!NOTE]
> The next queries are set in 90 days as an example, you can modify that value
> In both queries at the final appears the Local domain ".kazdemos.org" this is used to remove the domain from your device name, you can change by your own domain.

## This is the query used to collect data from MIP Scanner for discovery purpose

> [!IMPORTANT]
> Please remember that you need to have Scanner activities after the MPARR Collector was start to collect data, any previously activity is not collected.

```Kusto
AuditGeneral_CL
| where TimeGenerated >= now(-90d)
| where Common_ProcessName_s == "MSIP.Scanner" and Operation_s != "HeartBeat" and Operation_s != "FileDeleted" and Operation_s != "Search" and Operation_s != "Validate"
| extend Extensions = " "
| extend Filename = ObjectId_s
| extend FileName = replace_regex(Filename, @'^.*[\\/]', '')
| extend PATH = parse_path(ObjectId_s)
| parse PATH with * '"DirectoryPath":"' PATH1
| parse PATH1 with * '"AlternateDataStreamName":"' PATH2
| extend PATH2
| extend PATH = iff(PATH2 !contains 'http', PATH1 = trim_end(@'","DirectoryName":"."}',PATH1), PATH2 )
| extend Location = iff(PATH startswith "http", 'Web', Location = iff(PATH contains ":\\", 'Workstation', Location = iff(PATH contains "\\", 'Server', 'Unidentified')))
| extend PATH = iff(PATH startswith "http", replace_string(PATH,'//', '/'), replace_string(PATH, '\\\\', '\\'))
| project
Label = column_ifexists('SensitivityLabelEventData_SensitivityLabelId_g',''),
Date = column_ifexists('TimeGenerated',''),
ObjectId_s,
FileName,
PATH,
User = column_ifexists('UserId_s',''),
Activity = column_ifexists('Operation_s',''),
Location,
Version = column_ifexists('Common_ProductVersion_s',''),
IP = column_ifexists('ClientIP_s',''),
Extensions,
Workstation = column_ifexists('Common_DeviceName_s',''),
ResultStatus_s,
SensitiveInfoTypeData_s,
ProtectionEventData_IsProtected_b
| extend Workstation = trim_end(@'".kazdemos.org"',Workstation)
| extend Extensions = parse_path(FileName)
| parse Extensions with * '"Extension":"' Extensions
| extend Extensions = trim_end(@'","AlternateDataStreamName":""}',Extensions)
```

## This is the query used to collect data from MIP Scanner for discovery and labeling
> [!IMPORTANT]
> Please remember that you need to have Scanner activities after the MPARR Collector was start to collect data, any previously activity is not collected.
> These fields are not available if you are using MIP Scanner only for discovery purpose, but if you want to know about labels applies with the scanner you need use the next KQL
> Additional fields are:
> - ProtectionEventData_ProtectionType_s
> - ProtectionEventData_TemplateId_g
> - ProtectionEventData_ProtectionOwner_s

```Kusto
AuditGeneral_CL
    | where TimeGenerated >= now(-90d)
    | where Common_ProcessName_s == "MSIP.Scanner" and Operation_s != "HeartBeat" and Operation_s != "FileDeleted" and Operation_s != "Search" and Operation_s != "Validate"
    | extend Extensions = " "
    | extend Filename = ObjectId_s
    | extend FileName = replace_regex(Filename, @'^.*[\\\/]', '')
    | extend PATH = parse_path(ObjectId_s)
    | parse PATH with * '"DirectoryPath":"' PATH1
    | parse PATH1 with * '"AlternateDataStreamName":"' PATH2
    | extend PATH2
    | extend PATH = iff(PATH2 !contains 'http', PATH1 = trim_end(@'","DirectoryName":".*"}',PATH1), PATH2 )
    | extend Location = iff(PATH startswith "http", 'Web', Location = iff(PATH contains ":\\", 'Workstation', Location = iff(PATH contains "\\", 'Server', 'Unidentified')))
    | extend PATH = iff(PATH startswith "http", replace_string(PATH,'//', '/'), replace_string(PATH, '\\\\', '\\'))
    | project
        Label = column_ifexists('SensitivityLabelEventData_SensitivityLabelId_g',''),
        Date = column_ifexists('TimeGenerated',''),
        ObjectId_s,
        FileName,
        PATH,
        User = column_ifexists('UserId_s',''),
        Activity = column_ifexists('Operation_s',''),
        Location,
        Version = column_ifexists('Common_ProductVersion_s',''),
        IP = column_ifexists('ClientIP_s',''),
        Extensions,
        Workstation = column_ifexists('Common_DeviceName_s',''),
        ResultStatus_s,
        SensitiveInfoTypeData_s,
        ProtectionEventData_IsProtected_b,
        ProtectionEventData_ProtectionType_s,
        ProtectionEventData_TemplateId_g,
        ProtectionEventData_ProtectionOwner_s
    | extend Workstation = trim_end(@'".kazdemos.org"',Workstation)
    | extend Extensions = parse_path(FileName)
    | parse Extensions with * '"Extension":"' Extensions
    | extend Extensions = trim_end(@'","AlternateDataStreamName":""}',Extensions)
```
