> [!WARNING]
> PAGE UNDER CONSTRUCTION.

# The logic behind of MPARR_RMSData script and the access review over protected documents.

One of the most wanted and awaited reports from Microsoft Purview is the one related to Access Granted or Denied over protected documents.
And here it is:

![MPARR - External access](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/e9ea9927-9ff8-4a18-9cfc-b537017ae7f3)
<p align="center">MPARR - Access granted or denied over protected documents</p>

But, how to this kind information can be obtained?, because is not available under Microsoft Purview Audit(previously called Unified Audit Logs) or under Activity Explorer, happen that all the activities related to document protection are collected through the Azure RMS service and that information resides in another place.

To collect that information we need to use MPARR_RMSData.psq script, this one permit to consume the data from another API, Azure RMS API, and put that same information in the same workspace in Logs Analytics, and that permit to merge and integrate all the data collected by MPARR.

Next, we will explain how to this information is collected, and some considerations in this process.

## How to the information is collected

In the next diagram we will try to explain step by step how to the data is generated and collected through log services.

![MPARR - RMSData explained](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/002a4c6e-a714-43af-ba2f-6697098b0347)
<p align="center">MPARR-RMSData logs generated and collected</p>

1. When a user want to protect a document through a RMS Template or Sensitivity Label that user start an internal process to collect certificates, digital signatures protection templates and more, that is used to protect the information
1. Azure Right Management Service validate the connection and return all the previously mentioned items, recording in logs all these steps.
1. The document is protected and an attribute called Content ID is populated
![Get-AIPFileStatus](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/9da9acc6-5d3c-419a-9f7b-e47be063bb31)
<p align="center">Get-AIPFileStatus to show Content ID</p>
3. Document is sent or shared with protection
4. Recipient receive the document and the Microsoft 365 Apps for Enterprise(formerly called Office), or PDF reader applications, start a complete process in the background to request to Azure RMS to validate access en the permissions grantes.
5. Azure RMS validate the access and permissions granted to the requestor, returning grant or denied, and the rights over the document. All these steps are recorded on logs.
6. Document requestor receive a granted permission or denied access.

## What's next? 

Normally when you try to use tracking and revoke actions over protected documents, after you have DIRECT access to the file and obtain the Content ID, using Get-AIPFileStatus, you can request all the access through a PowerShell cmdlet, after you connect first to the AIP service, you can execute Get-AIPServiceTrackingLog with the Content ID and the result shown access denied or granted, rights granted, IP Address, requestor and some additional information.

![Get-AIPServiceTrackingLog](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/2dbe4072-03e1-4027-af9e-37175c602549)
<p align="center">Get-AIPServiceTrackingLog using Content ID and the data returned</p>

