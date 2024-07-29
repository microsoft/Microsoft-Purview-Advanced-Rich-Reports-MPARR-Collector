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
`Only Workloads' 
```KQL
AuditGeneral_CL
| where TimeGenerated >= now(-730d)
| summarize count() by Workload_s, Operation_s
| order by Workload_s asc
```
