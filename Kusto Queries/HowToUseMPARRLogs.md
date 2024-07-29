# How to use Office 365 Management Logs from Scratch with MPARR

![image](https://github.com/user-attachments/assets/86c2d538-8f0e-42e3-bec3-056cbfef81eb)

The tables that we will use for this exercise will be:
```
AuditAzureActiveDirectory_CL
AuditExchange_CL
AuditGeneral_CL
AuditSharePoint_CL
DLPAll_CL
```

Then depending on the volume of your data, we can get a certain period of time to avoid to many results and reduce costs.
> [!IMPORTANT]
> Each query executed in Logs Analytics have a cost, in general the cost is very small, nevertheless, extensive queries can generate a significant cost.

> [!NOTE]
> Because my demo environment is reduced and have a few activities, normally I use to get the past 2 years(past 730 days) of information, nevertheless, it's recommended reduce the query to a few days or hours, in that case toy can use the "Time range" option built-in in the Logs Analytics console.

```KQL
| where TimeGenerated >= now(-730d)
```

### Now we can start playing with the logs.

All the main logs coming from the Office 365 Management API and collected through the `MPARR_Collector2.ps1` script have two main columns:, that are:
- Operation_s
- Workload_s

> Operation_s: corresponds to all the different activities occurring in our environment, including end-user activities, administrator activities, and services running in the background. You can view a list of these activities at this [link](https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/blob/main/Support%20Information/MPARR%20-%20List%20of%20Operations%20Collected%20separated%20by%20table%20name.xlsx)
> "Workload_s: Each activity occurring in our Microsoft 365 tenant is grouped by service, referred to as 'workload' in the logs."

With the previous concepts on mind we can execute something like this:
```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| summarize count() by Workload_s, Operation_s
| order by Workload_s asc
```
This KQL returns all the Workloads currently available in our Logs, the operations related to each workload and the total amount of activities per operation.
![image](https://github.com/user-attachments/assets/0dbfe3c1-fd22-466d-b7b7-23285576fe8c)

Now, if you want to get a list of only operations or workloads we can execute in these ways:
> `Only Workloads`

```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| summarize count() by Workload_s
| order by Workload_s asc
```

> `Only Operations`

```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| summarize count() by Operation_s
| order by Operation_s asc
```

> `Only Workloads` Sample results

![image](https://github.com/user-attachments/assets/51bbd822-f803-46ac-84bd-b011b20f9286)

> `Only Operations` Sample results

![image](https://github.com/user-attachments/assets/b5ff309f-9089-4ad6-9059-a7544f8893f9)

> [!TIP]
> We can extend this exercise to all the tables in this way:
> ```KQL
> [Table Name]
> | where TimeGenerated >= now(-730d)
> | summarize count() by Workload_s, Operation_s
> | order by Workload_s asc
> ```

### Now we can go more in deep

Getting the list of workloads we can filter using this information, in this way:

```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| where Workload_s contains "MicrosoftTeams"
| summarize count() by Operation_s
| order by Operation_s asc
```
> Here we are collecting all the operations related to Microsoft Teams, at least all the activities that was occuring on my environment.

![image](https://github.com/user-attachments/assets/551906a9-bb69-436e-b614-00b39dcaafc7)

> [!IMPORTANT]
> We can see that the filter 'contains' is underlined with a red line. This is because KQL is case-sensitive. Using **'contains'** instead of **'=='** is more resource-intensive, but it allows for the identification of workloads that can be written in lowercase, uppercase or a mix. Nevertheless, if you are clear that the workload appears wrote in only one way, you can use **==**.

The next step is to obtain detailed information about a specific activity. Similarly, we can gather information about a specific workload, but this will include details on all the operations associated with that workload.

```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| where Operation_s contains "ChannelAdded"
```
> In the previous steps, we identified workloads and operations. In the last query, we obtained a list of Microsoft Teams operations, and now we are focusing on a specific operation called **'ChannelAdded'**. Based on my experience, operations with the same name are rarely found across different workloads, except for specific cases like **'DLPRuleMatch'** or **'SensitivityLabelApplied'**. Therefore, I did not include the workload filter in the query, but it can be added if needed.

![image](https://github.com/user-attachments/assets/412d5dc9-b99f-4fc9-b54c-76ffcc8319f3)

> [!TIP]
> Another field that comes into play is **'UserId_s'**, which can be used to identify activities or workloads associated with specific users or to narrow down the results to certain users.

```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| where Operation_s contains "ChannelAdded"
| where UserId_s contains "sebastian"
```
![image](https://github.com/user-attachments/assets/0bee0763-12c1-4813-9ce3-236978d4f0e8)

All this exercise permit to identify all the fields related to each activity, additionally, to the workload, operation and user.

### And then...

We can start to play with the different fields available on each operation or we can get all the activities per User.
```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| where UserId_s contains "mike"
| summarize count() by Workload_s, Operation_s
| order by Workload_s asc
```

> This previous query returns all activities per workload for the user named **'mike'**. In this environment, we have only one user named **Mike Wazowski**. The following image shows all activities related to Purview for this user in an administrator role.

![image](https://github.com/user-attachments/assets/ec978d81-9bef-4642-a727-747f36819fc8)
