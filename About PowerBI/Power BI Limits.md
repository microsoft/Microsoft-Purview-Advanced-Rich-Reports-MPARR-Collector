> [!WARNING]
> PAGE UNDER CONSTRUCTION.

# About Power BI Limits and the procedure to go beyond this limits.

Currently some limits are related to Power BI Pro on the past was a limit related to the maximum records returned by query, on the past that limit was related to 500k as a limit and today is related to size.
To avoid, or in certain way skip that limit, we use a special approach.

We will use the DLP case to explain this exercise and the step by step to achieve this.

The 1st step is get the dates, separated by Year, Month and Day, when we have certain kind of activities. To achieve this we will use a query like this one:
```Kusto
DLPAll_CL
| where TimeGenerated >= now(-730d)
| where Operation_s contains "DLPRule"
| summarize by 
    Year = datetime_part('Year',TimeGenerated), 
    Month = datetime_part('Month',TimeGenerated),
    Day = datetime_part('Day',TimeGenerated)
```
The previous KQL returns for all the data located on Logs Analytics, in my this case I set the retention period for 2 years, for that reason we are getting the data for the past 730 days. And the result from that query is each day when we have data vailable that match with the filter related to Operations, or Activities, that match with DLPRule.
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/c6874403-7eb8-4b27-8b6e-79d20d07f2f8"/></p>
<p align="center">KQL returning the dates where the filter match with activities.</p>

When this data is exported as an M Query for Power BI the result is this one:
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/1d0e51b6-24f3-4982-a137-63cb83282a73"/></p>
<p align="center">KQL as an M Query imported in Power BI.</p>

No we have the first step completed in this process.

The 2nd phase is a little more complex and require to take care on each steps that is required. To start in this procedure we will start with the specific query that we want to use to get our data, can be a simple query or a complex one, in this case we will use something very simple like the next one(we are not setting time, we are using default that is set for the past 24 hours):
```Kusto
DLPAll_CL 
| where Operation_s contains "DLP"
```

Is important to have at least one result to enable the option to export as an M Query, and then import on Power BI.
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/03c6e79e-6c1d-43be-9d42-a66f4bf2b05a"/></p>
<p align="center">KQL return on Logs Analytics.</p>

<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/1f2c8d6a-5b14-4464-91b7-d269b6ca9c84"/></p>
<p align="center">Power BI Query editor and the results from import KQL as an M query for DLP.</p>

Doing a right click over this Query we will find the option to create a function from this query, you will need to set a name for the function, you can use any kind of name.

<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/8c628294-c459-4c8d-bf7e-dac6c3e49e6c"/></p>
<p align="center">Create a function from a query.</p>

**In this point we need to select in the top menu the option Advanced Editor to edit our  new function**
_(is recommended to enable "Display Line Numbers" selecting at the top right corner menu)_
In the new function we will need to change the line 2 that contains this code:
```Kusto
let
    Source = () => let AnalyticsQuery =
```
By this one, is only add inside the parentheses this line _"Day as number, Month as number, Year as number"_:
```Kusto
let
    Source = (Day as number, Month as number, Year as number) => let AnalyticsQuery =
```

Then we will need to identify the line that contains the string timespan and remove that attribute with the value:
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/d5a2f4fc-2cae-4f3b-ac79-424fb46b106e"/></p>
<p align="center">Attribute to be removed.</p>

<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/f40acfa2-cacf-4d0b-93b9-fd14bfb3e4d5"/></p>
<p align="center">Query after attribute was removed.</p>

To continue some additional steps are needed on this configuration, we will need add the next query:
```Kusto
let querytime = todatetime(strcat("&Number.ToText(Month)&",'/',"&Number.ToText(Day)&",'/',"&Number.ToText(Year)&")); 
let BeginDay = startofday(querytime); 
let EndDay = endofday(querytime);
```
The previous code needs to be added starting on the line 4 just before the table name, in this case called DLPAll_CL
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/f72aa6bb-3a90-4df2-93e1-5d635be38a0c"/></p>
<p align="center">Time string query added to the function.</p>

And finally for this last step in the function we need to add the next query:
```Kusto
| where TimeGenerated between (BeginDay .. EndDay)
```
The final result is show in the image below.
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/a06312d9-6c27-4538-b964-e0dfa0cc9f7d"/></p>
<p align="center">Final step to configure the function on the advanced query editor.</p>

As a summary the steps are:
1. From the original query do a right click and select "Create Function" and add a name.
2. Press the Advanced Editor menu to modify the new function.
3. At the line 2 add the string _"Day as number, Month as number, Year as number"_ inside of the parentheses at source.
4. Search for the "timespan" attribute and remove with the value previously set.
5. At line 4 just before the Table Name, in this case DLPAll_CL, the string show previously.
6. Finally just after the Table Name add the string related to filter by Timegenerated.

After you finish the previous steps and you press the _"done"_ buttonyou will see the function in this way:
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/9b878be4-c65e-4cf0-81e9-9e2fd0311131"/></p>
<p align="center">Function interface.</p>

The last phase is call this function for each day on the first query created, that permit to call several times the query and permitting to go over the limit mentioned at the begins, doing several calls, one for each day identified with activities.

To do that we need to go to the first query and at the top menu select the tab called *"Add Column"* and the press *"Invoke Custom Function"* 
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/fdfc024a-ad55-4c04-b33e-b25ade8ead4e"/></p>
<p align="center">Add Column - Invoke Custom Function.</p>

In the pop-up window open you can set the name as default and then select the name of your function, after that a new window will be show and in this point, for each variable Day, Month and Year we need to select *"Column Name"* and each column name that match with the variable name, Day for Day, Month for Month, and Year for Year.
<p align="center">
<img src="https://github.com/microsoft/Microsoft-Purview-Advanced-Rich-Reports-MPARR-Collector/assets/44684110/e8506d2c-b9ac-44e4-b95b-2adb494dc592"/></p>
<p align="center">Selecting variables to call function.</p>

Those are all the steps required for this configuration, you can find about How to create reports from Scratch in our *YouTube Channel* at https://aka.ms/MPARR-YouTube, don't forget to subscribe, give us like to our videos, and follow us in our social networks.
