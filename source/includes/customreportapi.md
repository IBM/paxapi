# Custom Report API function

Custom Report functions can be used to interact with Custom Report worksheets. The Custom Report functions that are exposed through the IBM® Cognos® automation objects are:


## Create (Custom Report)

> Example

```vb
Public Sub Create()
    Reporting.CustomReports.create "http://computername", "Planning Sample", "plan_BudgetPlan", "Goal Input"
End Sub
```
Create generates an Custom Report based on the host system URL, server name, cube name, and view name.

### Syntax

The following string is the syntax for the Create method.

`Reporting.CustomReports.create "<host system URL>", "<server name>", "<cube name>", "<view name>"`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Custom Report is to be created from. | Alphanumeric string
server name | Name of the server which the Custom Report is to be created from. | Alphanumeric string
cube name | Name of the cube which the Custom Report is to be created from. | Alphanumeric string
view name | Name of the view which the Custom Report is to be created from. | Alphanumeric string

## CreatefromMDX (Custom Report)

> Example

```vb
Public Sub CreateFromMDX()
    Reporting.CustomReports.createfromMDX "http://vottepps06.canlab.ibm.com:9510/",
   "Planning Sample", "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].
   [Revenue]} ON 0, {[plan_time].[plan_time].[2004]} ON 1 FROM [plan_BudgetPlan]"
End Sub
```
CreateFromMDX generates a Custom Report based on a host system URL, server name, and MDX string.

### Syntax

The following string is the syntax for the CreatefromMDX method.

`Reporting.CustomReports.createfromMDX "<host system URL>", "<server name>", "<MDX statement>"`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Custom Report is to be created from. | Alphanumeric string
server name | Name of the server which the Custom Report is to be created from. | Alphanumeric string
MDX statement | MDX statement which the Custom Report is to be created from. | Alphanumeric string