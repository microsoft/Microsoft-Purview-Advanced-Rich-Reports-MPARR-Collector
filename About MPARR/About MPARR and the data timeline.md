# About MPARR Collector and MPARR RMSData scripts and the data timeline

Is very important the way to understand the Microsoft 365 Tenant records about activities when we talk about MPARR.

When you enable your Tenant for the first time, one of the oneshot configuration required is enable the audit activities, after we enable this we can see different kind of reports including Microsoft Purview Audit(previouslye called Unified Audit Logs) and Activity Explorer.

![MPARR - Timeline](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/d5940131-5442-4bb2-94cb-29f4219c1b08)
<p align="center">MPARR general overview and comparison about timeline</p>

All the activities from end users, administrators and services running in the background are collected, today we found more than 680 different activities and growing.
The normal consumption for all these activities is through Microsoft Purview Audit console, where we can create some kind of queries to find some specific information, the filters available shown some activities, but for a more detailed search you need to use PowerShell, and known those activities that we are looking for. Depending in your licensing the information collected can be the last 180 days (after the new updated) for E3 and 365 days for E5, only considering that, no matter if you are using Web interface or PowerShell, you can export 10k records per query, this limit will be increased.

The other option, to collect an reduced segment of all this data, is through Activity Explorer, in this case the information collected is related to end user activities for Microsoft Information. In this case we can have activities related to DLP, MIP, Retention and others, always related to end users; from this option we can export from the Web Interface 10k records or 30K records using PowerShell, but only accessing to the last 30 days. In this case the limits will be increased, but in some cases is still a small number.

<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/f2b03211-a5dd-476a-8d29-6440915576ed"></p>
<p align="center">Microsoft 365 & Microsoft Purview current log architecture</p>

One of the most common options to collect all the records without limits is use the Office 365 Management API, normally used by SIEM solutions, but this API is thought to collect the daily information, and in that order of ideas, only the past 7 days can be requested.
MPARR uses the Office 365 management API and based in that architecture the Collector collects the daily information. For some specific cases MPARR_Collector.ps1 have some attributes that can be used to collect past data but only if that activity occur in the past 7 days. (You can see the first diagram that explains)

Understanding the importance of collect data for a bigger time frame, MPARR can send the data to Logs Analytics by default, where the data can be collected for time frame of 2 years and have the capability to all the information or certain information beed retained for 7 years.

With the data stored in a Logs Analytics workspace we can consume the information through Sentinel or Power BI, this permit to access to all the activities generated on Microsoft 365 and Microsoft Purview and easily identify some specific actions. With the supporting scripts, MPARR permit to filter the data using Microsoft Entra(previously called Azure AD) attributes enriching the data and permiting to generate more advanced reports.
