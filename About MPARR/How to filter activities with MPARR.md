## How to filter operations with MPARR

To filter certain operations with MPARR, consider the following concepts and files:

#### Overview

According to the [Office 365 Management API documentation](https://learn.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference#working-with-the-office-365-management-activity-api), the Office 365 Management API collects logs grouped into five categories:

- Audit Azure Active Directory
- Audit Exchange
- Audit General
- Audit SharePoint
- DLP All

These groups correspond to tables created in Logs Analytics..

#### Managing Log Download Options

To manage the capability of turning on or off the option to download logs from these sources, you can use the `schemas.json` file. In this file, you can easily change the values for each group from "True" to "False" or vice versa.

Additionally, this file contain the capability to use filtering capabilities available in the main script "MPARR_Collector2.ps1" that can be set to "Contains" or "NotContains" based on Operations. You can find a list of Operations through this [link](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/blob/main/Support%20Information/MPARR%20-%20List%20of%20Operations%20Collected%20separated%20by%20table%20name.xlsx)

#### Filtering Operations

The `schemas.json` file also supports filtering capabilities, which can be utilized in the main script `MPARR_Collector2.ps1`. You can set filters to "Contains" or "NotContains" based on specific operations. A list of operations can be found [here](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/blob/main/Support Information/MPARR - List of Operations Collected separated by table name.xlsx).

#### Schemas.json file

```json
{
  "Audit.AzureActiveDirectory": "True",
  "FilterAuditAzureActiveDirectory": "NotContains",
  "Audit.Exchange": "True",
  "FilterAuditExchange": "Contains",
  "Audit.SharePoint": "True",
  "FilterAuditSharePoint": "Contains",
  "Audit.General": "True",
  "FilterAuditGeneral": "Contains",
  "DLP.All": "True",
  "FilterDLPAll": "Contains"
}
```

#### Customizing Activity Collection

The API can collect more than 750 different activities related to various services like Purview, Forms, Project, Planner, Power BI, Power Apps, Dynamics, Streams, and Viva, among others. Some organizations may not need reports for all these services and may focus on specific activities related to Data Loss Prevention (DLP), Sensitivity labels, or Retention labels. In such cases, the filtering capabilities available through the `MPARR_Collector2.ps1` script can be used to collect only the relevant data for that purpose.

As an example we can use the script in this way:

```powershell
.\MPARR_Collector2.ps1 -FilterAuditExchange "DLP|Sensitivity" -FilterAuditSharePoint "DLP|Sensitivity" -FilterAuditGeneral "DLP|Sensitivity" -FilterDLPAll "DLP|Sensitivity"
```

In the previous execution we are getting all the activities from Exchange, SharePoint, General and DLP groups from the API that `contains` activities that match on the string name with DLP or Sensitivity, in that case DLPRuleMatch, DLPRuleUndo, Get-DLPPolicy, SensitivityLabelApllied, SensitivityLabelRemove, and several other activities.

#### Task Scheduler and PowerShell odd behavior

Nevertheless, execute PowerShell scripts using attributes through Task Scheduler can generate some odd behaviors, for that same reason the best alternative is use a new script called "[MPARR2_run_me.ps1](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/blob/main/MPARR2/MPARR2_run_me.ps1)" that call "MPARR_Collector2.ps1" adding the attributes related to the filtering capabilities. And replacing under task scheduler the call to the main script for the MPARR2_run_me one. 

![image](https://github.com/user-attachments/assets/0b3d50db-04be-4c22-8931-3f7c1c464487)
