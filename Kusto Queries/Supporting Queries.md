# Supporting elements used in Power BI Templates

> [!NOTE]
> The next queries are set in 30 days as an example, you can modify that value

## These are queries over AzureADUsers_CL table

> [!IMPORTANT]
> Please remember that you can change the attributes collected from Microsoft Entra ID(Azure AD) modifying the script MPARR-AzureADUsers.ps1

### Query used to collect records from AzureADUsers_CL table removing duplicated data
```Kusto
AzureADUsers_CL 
| where TimeGenerated >= now(-30d)
| summarize arg_max(TimeGenerated, UserPrincipalName_s, DisplayName_s, AssignedLicenses_s, City_s, JobTitle_s, Department_s, Mail_s, OfficeLocation_s, UserID_g) by UserPrincipalName_s
```

### Query used to collect records from AzureADDomains_CL table removing duplicated data
```Kusto
AzureADDomains_CL
| where TimeGenerated >= now(-30d)
| summarize arg_max(TimeGenerated, Domain_s) by Domain_s
```
