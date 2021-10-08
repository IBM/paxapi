# Dynamic Report API functions

Dynamic Report functions can be used to interact with Dynamic Report worksheets. 

Dynamic Report functions can use the following PropertyAccessor objects:

PropertyAccessor | Description | 
--------- | ------- | -----------
GetReports() | Gets the collection of Dynamic Report objects from the active book.
GetAt(sheet) | Gets the collection of Dynamic Report objects from the specified sheet name in the active book.
Get(ignored, sheet, id) | Gets the Dynamic Report object from the specified sheet name, with the given ID, in the active book.

The Dynamic Report functions that are exposed through the IBM® Cognos® automation objects are:

## Create (Dynamic Report)

> Example

```vb
Public Sub Create()
    Reporting.DynamicReports.create "http://computername", "Planning Sample", 
    "plan_BudgetPlan", "Goal Input"
End Sub
```
Create generates an Dynamic Report based on the host system URL, server name, cube name, and view name.

### Syntax

The following string is the syntax for the Create method.

`Reporting.DynamicReports.create "<host system URL>", "<server name>", "<cube name>", "<view name>" `

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Dynamic Report is to be created from. | Alphanumeric string
server name | Name of the server which the Dynamic Report is to be created from. | Alphanumeric string
cube name | Name of the cube which the Dynamic Report is to be created from. | Alphanumeric string
view name | Name of the view which the Dynamic Report is to be created from. | Alphanumeric string

## CreatefromMDX (Dynamic Report)

> Example

```vb
Public Sub CreateFromMDX()
    Reporting.DynamicReports.CreateFromMDX "http://vottepps06.canlab.ibm.com:9510/",
   "Planning Sample", "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].
   [Revenue]} ON 0, {[plan_time].[plan_time].[2004]} ON 1 FROM [plan_BudgetPlan]"
End Sub
```
CreateFromMDX generates a Dynamic Report based on a host system URL, server name, and MDX string.

### Syntax

The following string is the syntax for the CreatefromMDX method.

`Reporting.DynamicReports.CreatefromMDX "<host system URL>", "<server name>", "<MDX statement>"`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Dynamic Report is to be created from. | Alphanumeric string
server name | Name of the server which the Dynamic Report is to be created from. | Alphanumeric string
MDX statement | MDX statement which the Dynamic Report is to be created from. | Alphanumeric string

## GetMDX

> Example

```vb
MsgBox Reporting.DynamicReports.GetAt(Application.ActiveSheet.name).Item(0).GetMDX
```
This API call is used to return the MDX for theDynamic Report row.

### Syntax

The following string is the syntax for the GetMDX method.

`Reporting.DynamicReports.GetAt(Application.ActiveSheet.name).Item(0).GetMDX`

## FormatAreaVisible

> Example

```vb
Public Sub Create()
    Reporting.DynamicReports.GetAt(Application.ActiveSheet.name).Item(0).FormatAreaVisible (true)
End Sub
```
This API call is used to show and hide the formatting area in a Dynamic Report.

### Syntax

The following string is the syntax for the FormatAreaVisible method.

`Reporting.DynamicReports.GetAt(Application.ActiveSheet.name).Item(0).FormatAreaVisible (<true/false>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
true/false | true: Shows the formatting area in the Dynamic Report. false: Hides the formatting area in the Dynamic Report. | Boolean

## Refresh (Dynamic Report)

> Example

```vb
Reporting.DynamicReports.GetAt(DynamicReports.Worksheet.Name).Item(0).Refresh
```
This API call is used to refresh a Dynamic Report.

### Syntax

The following string is the syntax for the Refresh method.

`Reporting.DynamicReports.GetAt().Item(<Dynamic Report ID>).Refresh`

<aside class="notice">
The `Dynamic Report ID` is not the actual ID, but the ordinal position. If you have multiple Dynamic Reports on a sheet, you can use `GetAt(#)` to define the ordinal position of the Dynamic Report. <br> <b>Examples:</b> <br> Reporting.DynamicReports.GetAt().Item(0).Rebuild  refers to the first Dynamic Report. <br>
Reporting.DynamicReports.GetAt().Item(1).Rebuild refers to the second Dynamic Report.
</aside>

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Dynamic Report ID | The ID of the Dynamic Report that is to be refreshed. | Integer


## Rebuild (Dynamic Report)

> Example

```vb
Reporting.DynamicReports.GetAt(ActiveCell.Worksheet.Name).Item(0).Rebuild
```
This API call is used to rebuild a Dynamic Report.

### Syntax

The following string is the syntax for the Rebuild method.

`Reporting.DynamicReports.GetAt().Item(<Dynamic Report ID>).Rebuild`

<aside class="notice">
The `Dynamic Report ID` is not the actual ID, but the ordinal position. If you have multiple Dynamic Reports on a sheet, you can use `GetAt(#)` to define the ordinal position of the Dynamic Report. <br> <b>Examples:</b> <br> Reporting.DynamicReports.GetAt().Item(0).Rebuild  refers to the first Dynamic Report. <br>
Reporting.DynamicReports.GetAt().Item(1).Rebuild refers to the second Dynamic Report.
</aside>

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Dynamic Report ID | The ID of the Dynamic Report that is to be rebuilt. | Integer


## RebuildActiveSheet

> Example

```vb
Public Sub RebuildMyDynamicReportSheet()
    Dim test As Object
    Set test = Reporting
    'New call to rebuild active sheet.
    test.DynamicReports.RebuildActiveSheet
End Sub
```
This API call is used to rebuild the active sheet.

### Syntax

The following string is the syntax for the RebuildActiveSheet method.
`Reporting.DynamicReports.RebuildActiveSheet`

## RebuildActiveBook

> Example

```vb
Public Sub RebuildMyDynamicReportWorkbook()
    Dim test As Object
    Set test = Reporting
    'New call to rebuild active workbook.
    test.DynamicReports.RebuildActiveWorkbook
End Sub
```

This API call is used to rebuild the Dynamic Reports in the active workbook, even if the Dynamic Reports are on different sheets.

### Syntax

The following string is the syntax for the Rebuildbook method.
`Reporting.DynamicReports.RebuildActiveWorkbook`

