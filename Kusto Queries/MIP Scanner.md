# Kusto Query for MIP Scanner

## This is the query used to collect data from MIP Scanner

<br/>Please remember that you need to have Scanner activities after the MPARR Collector was start to collect data, any previously activity is not collected.

<br/>AuditGeneral_CL
<br/>    | where TimeGenerated >= now(-90d)
<br/>    | where Common_ProcessName_s == "MSIP.Scanner" and Operation_s != "HeartBeat" and Operation_s != "FileDeleted" and Operation_s != "Search" and Operation_s != "Validate"
<br/>    | extend Extensions = " "
<br/>    | extend Filename = ObjectId_s
<br/>    | extend FileName = replace_regex(Filename, @'^.*[\\\\\/]', '')
<br/>    | extend PATH = parse_path(ObjectId_s)
<br/>    | parse PATH with * '"DirectoryPath":"' PATH1
<br/>    | parse PATH1 with * '"AlternateDataStreamName":"' PATH2
<br/>    | extend PATH2
<br/>    | extend PATH = iff(PATH2 !contains 'http', PATH1 = trim_end(@'","DirectoryName":".*"}',PATH1), PATH2 )
<br/>    | extend Location = iff(PATH startswith "http", 'Web', Location = iff(PATH contains ":\\\\", 'Workstation', Location = iff(PATH contains "\\\\", 'Server', 'Unidentified')))
<br/>    | extend PATH = iff(PATH startswith "http", replace_string(PATH,'//', '/'), replace_string(PATH, '\\\\\\\\', '\\\\'))
<br/>    | project
<br/>        Label = column_ifexists('SensitivityLabelEventData_SensitivityLabelId_g',''),
<br/>        Date = column_ifexists('TimeGenerated',''),
<br/>        ObjectId_s,
<br/>        FileName,
<br/>        PATH,
<br/>        User = column_ifexists('UserId_s',''),
<br/>        Activity = column_ifexists('Operation_s',''),
<br/>        Location,
<br/>        Version = column_ifexists('Common_ProductVersion_s',''),
<br/>        IP = column_ifexists('ClientIP_s',''),
<br/>        Extensions,
<br/>        Workstation = column_ifexists('Common_DeviceName_s',''),
<br/>        ResultStatus_s,
<br/>        SensitiveInfoTypeData_s,
<br/>        ProtectionEventData_IsProtected_b
<br/>    | extend Workstation = trim_end(@'"& LocalDomain &"',Workstation)
<br/>    | extend Extensions = parse_path(FileName)
<br/>    | parse Extensions with * '"Extension":"' Extensions
<br/>    | extend Extensions = trim_end(@'","AlternateDataStreamName":""}',Extensions)
