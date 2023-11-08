> [!WARNING]
> PAGE UNDER CONSTRUCTION.

# The logic behind of MPARR_RMSData script and the access review over protected documents.

One of the most wanted and awaited reports from Microsoft Purview is the one related to Access Granted or Denied over protected documents.
And here it is:

![MPARR - External access](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/e9ea9927-9ff8-4a18-9cfc-b537017ae7f3)
<p align="center">MPARR - Access granted or denied over protected documents</p>

But, how to this kind information can be obtained?, because is not available under Microsoft Purview Audit(previously called Unified Audit Logs) or under Activity Explorer, happen that all the activities related to document protection are collected through the Azure RMS service and that information resides in another place.

To collect that information we need to use MPARR_RMSData.ps1 script, this one permit to consume the data from another API, Azure RMS API, and put that same information in the same workspace in Logs Analytics, and that permit to merge and integrate all the data collected by MPARR.

Next, we will explain how to this information is collected, and some considerations in this process.

## How to the information is collected

In the next diagram we will try to explain step by step how to the data is generated and collected through log services.
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/002a4c6e-a714-43af-ba2f-6697098b0347"/></p>
<p align="center">MPARR-RMSData logs generated and collected</p>

1. When a user want to protect a document through a RMS Template or Sensitivity Label that user start an internal process to collect certificates, digital signatures protection templates and more, that is used to protect the information
1. Azure Right Management Service validate the connection and return all the previously mentioned items, recording in logs all these steps.
1. The document is protected and an attribute called Content ID is populated
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/9da9acc6-5d3c-419a-9f7b-e47be063bb31"/></p>
<p align="center">Get-AIPFileStatus to show Content ID</p>
3. Document is sent or shared with protection
4. Recipient receive the document and the Microsoft 365 Apps for Enterprise(formerly called Office), or PDF reader applications, start a complete process in the background to request to Azure RMS to validate access en the permissions grantes.
5. Azure RMS validate the access and permissions granted to the requestor, returning grant or denied, and the rights over the document. All these steps are recorded on logs.
6. Document requestor receive a granted permission or denied access.

## What's next? 

Normally when you try to use tracking and revoke actions over protected documents, after you have DIRECT access to the file and obtain the Content ID, using Get-AIPFileStatus, you can request all the access through a PowerShell cmdlet, after you connect first to the AIP service, you can execute Get-AIPServiceTrackingLog with the Content ID and the result shown access denied or granted, rights granted, IP Address, requestor and some additional information.

<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/2dbe4072-03e1-4027-af9e-37175c602549"/></p>
<p align="center">Get-AIPServiceTrackingLog using Content ID and the data returned</p>

Do the previous steps for each file inside of our organization is almost impossible, here is when MPARR-RMSData.ps1 do their magic.
MPARR-RMSData run the cmdlet Get-AipServiceUserLog tha collects all the information generated in the points 2 and 6 from the previous Diagram and send all that information to the RMSData_CL table in Logs Analytics and everytime time that the field "Content ID" appears with information, the same script execute the cmdlet Get-AIPServiceTrackingLog for that value and the data is collected in the RMSDataDetails_CL table.
This exercise permit to have a complete track every time that a protected document is tried to open and the results from that action, and collect all that information.

> The big deal... this information doesn't contains file name, or path of the file, or other relevant information.

Here coming the relationship between MPARR-RMSData and MPARR_Collector, the activities from the end user in the point 1 are sent to Microsoft 365 Audit logs that can be collected by our Collector and are sent to, in this case, the table AuditGeneral_CL to the workspace in Logs Analytics. [KQL for Audit General used in this report](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/blob/main/Kusto%20Queries/Common%20KQL%20used%20on%20Power%20BI.md#mparr_collectorps1)

Happens as we mentiones previously that MPARR_Collector today collects more than 680 different activities, or at least the activities that we was identifying until now, that is too many information and we need to reduce the data request from Logs Analytics, in that case the previous Query permit to collect only the information needed for this report.

And here we can see the relationship between all the tables in the Power BI template:
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/af990575-92b8-4a4a-9332-3b9c3fc6d6be"/></p>
<p align="center">MPARR - Access review data relationship</p>

We found that we can create a relationship between:
- RMSDataDetails_CL : ContentId_g -> RMSData_CL : content_id_g
- RMSData_CL : template_id_g -> AuditGeneral_CL : CurrentProtectionType_templateId_g

In this last point we need to understand that MPARR can show only the information collected, and if you review the point related to [About MPARR and the data timeline](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/blob/main/About%20MPARR/About%20MPARR%20and%20the%20data%20timeline.md#about-mparr-collector-and-mparr-rmsdata-scripts-and-the-data-timeline); we need to understand that any documented protected previously, to starting use MPARR, doesn't appears on the Logs, for that reason at the begins can be common have reports about access denied or granted without identify the file name and the path.
