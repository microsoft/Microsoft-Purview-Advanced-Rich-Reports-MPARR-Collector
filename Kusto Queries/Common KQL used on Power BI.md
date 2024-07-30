> [!WARNING]
> PAGE UNDER CONSTRUCTION.

# Common Kusto Queries used in Power BI templates

## Queries to generate a table with day, month and year when operations appears
Historicaly you can find some information about Power BI limits related to the amount of data that can be downloaded in a single query, a way to avoid that limit is create a function and merge with the tables created in this section, that permit create a query call per each day.

### MPARR-RMSData.ps1 (Azure RMS API)
> [!NOTE]
> MPARR-RMSData.ps1 generate 2 tables the principal one, RMSData_CL, created in base to all the activities related to protect or unprotect documents and emails, and operations related. When a file or document is protected or someone tries to access, that information is registered at RMSDataDetails_CL this table contains all the activities related to access denied or granted over a protected document. 

```Kusto
RMSDataDetails_CL 
| where TimeGenerated > now(-730d)
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```

```Kusto
RMSData_CL 
| where TimeGenerated > now(-730d)
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```

### MPARR-Collector.ps1 (Office 365 Management API)
Exist some documentation about some limits on Power BI when we try to download certain number of information using a query, with these queries we can obtain the dates where we can found activities in each table, tha dates are separated in Year, Month and Day, after that we can use other queries as a functions and, finally, generate a query per day increasing the number of results, a kind of way to skip the limit.

```Kusto
AuditGeneral_CL  
| where TimeGenerated > now(-730d)
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```

```Kusto
AuditExchange_CL 
| where TimeGenerated > now(-730d)
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```

```Kusto
AuditSharePoint_CL 
| where TimeGenerated > now(-730d)
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```

```Kusto
DLPAll_CL 
| where TimeGenerated > now(-730d)
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```

## Detailed information collected

### MPARR-RMSData.ps1
```Kusto
RMSDataDetails_CL 
| where TimeGenerated >= now(-30d)
| project 
    TimeGenerated,
    ContentId_g,
    Issuer_s,
    RequestTime_t,
    RequesterType_s,
    RequesterEmail_s,
    RequesterDisplayName_s,
    RequesterLocation_IP_s,
    Rights_s,
    Successful_b
```

```Kusto
RMSData_CL 
| where TimeGenerated > now(-730d)
| where content_id_g != "" and template_id_g != "" and user_id_s !contains "@aadrm.com"
| project 
    TimeGenerated,
    content_id_g,
    template_id_g,
    date_s,
    time_s,
    request_type_s,
    user_id_s,
    result_s,
    owner_email_s,
    issuer_s,
    date_published_s,
    c_info_s,
    c_ip_s
```

### MPARR_Collector.ps1

#### Query used in "MPARR - MIP Access overview" Power BI template

```Kusto
AuditGeneral_CL
| where CurrentProtectionType_templateId_g!= ""
| project 
    TimeGenerated,
    SensitivityLabelEventData_JustificationText_s,
    PreviousProtectionType_templateId_g,
    CurrentProtectionType_templateId_g,
    IrmContentId_g,
    Platform_s,
    ContentType_s,
    CurrentProtectionType_documentEncrypted_b,
    CurrentProtectionType_owner_s,
    CurrentProtectionType_protectionType_d,
    ProtectionEventType_d,
    TargetLocation_d,
    SensitivityLabelEventData_ActionSource_d,
    SensitivityLabelEventData_LabelEventType_d,
    SensitivityLabelEventData_SensitivityLabelId_g,
    LabelId_g,
    UserId_s,
    Id_g,
    RecordType_d,
    Operation_s,
    Workload_s,
    EventCreationTime_t,
    ClientIP_s,
    ObjectId_s,
    Application_s,
    DeviceName_s
```
