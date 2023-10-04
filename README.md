# Project

## Welcome to Microsoft Purview Advanced Rich Reports (MPARR) Collector.

Today have the right information in the right moment can be a great business value, we are talking about implement security and compliance, one of the most importance things are understanding that this accomplishment is a business goal, in that order of ideas, to have reports, business-friendly, to see how our end users are using and adopting these tools is a global benefit. In that order of ideas, this solution take the information available under the Microsoft 365 services and give the capabilities to present this information to different business units, given the capacity to c-level users have access to business metrics related to compliance.

![MPARR - Solution Achitecture](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/8b324ae4-27bf-4e38-9aaf-cc19cfee27a3)
<p align="center">Current Architecture for MPARR</p>

Today one of the principal challenges in all organizations is stay align with the Compliance principles, each organization define their own priorities, and policies definitions. But, in all the cases they need to involve the complete organization, and to involve we need to show the right information at the right time.
Office 365 Management API collect all the information available through Unified Auditing tool, this helps to Security, Compliance and IT areas looking for some specific information and generate some reports but is not possible easily to show that information to different business units, and they don’t have the time neither to prepare more detailed reports.

![MPARR - Solution dependencies](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/d2a303e1-2333-415a-ba0c-aae1b06972fb)
<p align="center">Variables needed to set in laconfig.json files and the TableNames created on Logs Analytics</p>
 
In that order of ideas, the solution presented next permit to have a robust solution to collect all the data and prepare reports with specific scopes to specific audiences, without require special permissions or additional knowledge to understand the security tools.
This solution collects all the information available through Office 365 Management API and store this information on Logs Analytics workspace, this one can be the same used for Sentinel (we will discuss more this point next), from this workspace the information can be consumed using Power BI desktop to create advanced rich reports to publish then with Power BI online workspaces, this step permit to generate different workspaces for different audiences. To give more added value to these reports, some additional scrips are delivered, to collect as example the data related to Azure AD attributes, this one permit to create reports based on location, country, business units and any other Azure AD attribute available.

![MPARR - Solution data in Logs Analytics](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/356c5f6a-8e31-40af-9b1a-398c1e4df4d0)
<p align="center">TableNames created on Logs Analytics and use</p>
 
As we said previously, because this information can be stored on the same workspace used for Sentinel, this information can be utilized to generate workbooks with more detailed information for Security monitoring.
In this article, we will see how we can implement this script to start collecting the information and consume that information.

Some Power BI reports that can be created:

![MPARR - DLP overview](https://user-images.githubusercontent.com/44684110/215560498-5438d724-baf1-4a03-aea2-3cce1d7a88c5.png)
<p align="center">DLP Overview, department filter and dates filter</p>

![MP8851~1](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/69d18aea-6549-47e6-9491-693d9c2a2f6d)
<p align="center">Access denied and grantes over protected documents</p>

![MPARR - Unified Labeling Overview](https://user-images.githubusercontent.com/44684110/215561587-7e9507a7-b6b6-46a0-b950-abd7450bc2a0.png)
<p align="center">Unified Labeling Overview filtering by Deparment and Country</p>

![MPARR - Worldwide Operations review](https://user-images.githubusercontent.com/44684110/215561913-5dd632d1-fdbc-4eaf-8a36-13c3c5773791.png)
<p align="center">Worlwdide activities filtering by Operation</p>

![MPARR - MIP Scanner](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/53fe5694-fbcd-4cf0-9cab-4a2cb055a633)
<p align="center">MIP Scanner dashboard</p>

## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
