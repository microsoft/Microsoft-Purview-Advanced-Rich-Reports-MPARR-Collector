# Supporting elements used in Power BI Templates

> [!NOTE]
> The next queries are set in 30 days as an example, you can modify that value

## These are queries over tables created using Microsoft Graph API

The next information can be collected through the use of these scripts:
- MPARR-AzureADUsers
- MPARR-AzureADDomains
- MPARR-AzureADRoles

All these scripts are connected to Microsoft Graph API through the permissions added under the Microsoft Entra(previously called Azure AD) App, normally called MPARR-Datacollector. And this scripts uses a Current User certificate, if yu need a Local Machine certificate some little changes can be do it for that.

### Query used to collect records from AzureADUsers_CL table removing duplicated data
> [!NOTE]
> To obtain values you need previously executed MPARR-AzureADUsers.ps1 script

> [!IMPORTANT]
> Please remember that you can change the attributes collected from Microsoft Entra ID(Azure AD) modifying the script MPARR-AzureADUsers.ps1

```Kusto
AzureADUsers_CL 
| where TimeGenerated >= now(-30d)
| project 
    TimeGenerated = column_ifexists('TimeGenerated',''),
    UserPrincipalName_s = column_ifexists('UserPrincipalName_s',''),
    DisplayName_s = column_ifexists('DisplayName_s',''),
    AssignedLicenses_s = column_ifexists('AssignedLicenses_s',''),
    City_s = column_ifexists('City_s',''),
    JobTitle_s = column_ifexists('JobTitle_s',''),
    Department_s = column_ifexists('Department_s',''),
    Mail_s = column_ifexists('Mail_s',''),
    OfficeLocation_s = column_ifexists('OfficeLocation_s',''),
    UserID_g = column_ifexists('UserID_g',''),
    LastAccess_t = column_ifexists('LastAccess_t','')
| summarize arg_max(TimeGenerated, UserPrincipalName_s, DisplayName_s, AssignedLicenses_s, City_s, JobTitle_s, Department_s, Mail_s, OfficeLocation_s, UserID_g, LastAccess_t) by UserPrincipalName_s
```

### Query used to collect records from AzureADDomains_CL table removing duplicated data
```Kusto
AzureADDomains_CL
| where TimeGenerated >= now(-30d)
| summarize arg_max(TimeGenerated, Type) by Domain_s
```

### Query used to collect records from AzureADRoles_CL table removing duplicated data
```Kusto
AzureADRoles_CL 
| where TimeGenerated >= now(-30d)
| summarize arg_max(TimeGenerated, Description_s, Members_s, RoleID_g) by DisplayName_s
```

### Query used to collect records from MSProducts_CL table removing duplicated data
> [!NOTE]
> This script doesn't use any kind of API, only takes a CSV file with all the Microsoft licensing friendly-name and create a matrix in Logs Analytics with the IDs, short names and knowed names

```Kusto
MSProducts_CL 
| where TimeGenerated >= now(-30d)
| summarize arg_max(Product_Display_Name_s, TimeGenerated, String_Id_s, Service_Plan_Name_s, Service_Plans_Included_Friendly_Names_s) by GUID_g, Service_Plan_Id_g
```

## These are queries over tables created using Exchange Online Manage API

The next information can be collected through the use of these scripts:
- MPARR-SITData
- MPARR-LabelData

### Query used to collect records from Labels_CL table removing duplicated data
```Kusto
Labels_CL
| where TimeGenerated >= now(-30d)
| summarize arg_max(TimeGenerated, Name_s, Guid_g, Priority_d, ParentLabelDisplayName_s  ) by DisplayName_s
```

### Query used to collect records from SITs_CL table removing duplicated data
```Kusto
SITs_CL 
| where TimeGenerated >= now(-30d)
| summarize arg_max(TimeGenerated, Name_s, Publisher_s, Type_s, RecommendedConfidence_d, Description_s) by SIT_Id_g
```

