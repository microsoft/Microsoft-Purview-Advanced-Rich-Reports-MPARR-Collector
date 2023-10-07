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
