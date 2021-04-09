# Quick Report API functions

Quick Report functions can be used to interact with Quick Report worksheets. 

Quick Report functions can use the following PropertyAccessor objects:

PropertyAccessor | Description | 
--------- | ------- | -----------
Count | Counts the number of Quick Reports in the active book.
GetReports() | Gets the collection of Quick Report objects from the active book.

The Quick Report functions that are exposed through the IBM® Cognos® automation objects are:

## Clear (Quick Report)

> Example

```vb
Public Sub Clear()
    Reporting.GetCurrentReport(<ActiveCell>).Clear
End Sub
```
Clear is used to clear data from the Quick Report.

### Syntax

The following string is the syntax for the Clear method.

`Reporting.GetCurrentReport(<ActiveCell>).Clear`

## ColumnHierarchies

> Example

```vb
Sub ColumnHierarchies()
    Dim columns As String
    For Each Column In cafe.QuickReports.Get("0").ColumnDimensions
        If columns <> "" Then
            columns = columns & ", " & vbNewLine
        End If
        columns = columns & Column
    MsgBox "Columns:" columns
End Sub
```
ColumnHierarchies is used to return the hierarchies that exist in the columns of a Quick Report report.

### Syntax

The following string is the syntax for the ColumnHierarchies method.

`cafe.QuickReports.Get("<Quick Report ID").ColumnDimensions`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Quick Report ID | The ID of the Quick Report that the column hierarchies are being returned from | Integer

## Commit

> Example

```vb
Public Sub Commit()
    Reporting.GetCurrentReport(<ActiveCell>).Commit True
End Sub
```
Commit is used to commit the Quick Report report.

### Syntax

The following string is the syntax for the Commit method.

`Reporting.GetCurrentReport(<ActiveCell>).Commit <True>`

## Create (Quick Report)

> Example

```vb
Public Sub Create()
    Reporting.QuickReports.Create "http://computername/", "Planning Sample", 
    "plan_BudgetPlan", "Goal Input"
End Sub
```
Create generates a Quick Report based on the host system URL, server name, cube name, and view name.

### Syntax

The following string is the syntax for the Create method.

`Reporting.QuickReports.Create "<host system URL>", "<server name>", "<cube name>", "<view name>" `

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Quick Report is to be created from. | Alphanumeric string
server name | Name of the server which the Quick Report is to be created from. | Alphanumeric string
cube name | Name of the cube which the Quick Report is to be created from. | Alphanumeric string
view name | Name of the view which the Quick Report is to be created from. | Alphanumeric string

## CreateFromCVS (Quick Report)

> Example of the syntax for updating the common view specification of a report:

```vb
Reporting.QuickReports.CreateFromCVS("http://server-example.ibm.com", "Planning Sample", 
{
  "MDX": "SELECT {([d1].[h1].[line 2],[d3].[h1].[2004]),([d1].[h1].[line 2],[d3].[h1].[Q1-2004]),([d1].[h1].[line 2],[d3].[h1].[Jan-2004])}  DIMENSION PROPERTIES MEMBER_UNIQUE_NAME, LEVEL_NUMBER, CHILDREN_CARDINALITY ON 0  FROM [my_Cube] WHERE ( [d2].[h1].[toys], [d4].[h1].[USD], [d5].[h1].[Sales] )  CELL PROPERTIES CELL_ORDINAL, VALUE, FORMATTED_VALUE, FORMAT_STRING, UPDATEABLE, TM1UPDATEABLE, ANNOTATED, CONSOLIDATED",
  "Meta": {
    "Aliases": {
      "[d1].[h1]": "english",
      "[d3].[h1]": "english",
      "[d2].[h1]": "SKU"
    },
    "ExpandAboves": {
      "[d1].[h1]": false,
      "[d1].[h2]": true,
      "[d2].[h1]": false
    },
    "ContextSets": {
      "[d2].[h1]": {
        "Expression": "{ HIERARCHIZE( { TM1SUBSETALL([d2]) } ) }"
      },
      "[d4].[h1]": {
        "SubsetName": "Default"
      },
      "[d5].[h1]": {
        "SubsetName": "All Deparments",
        "IsPublic": true
      }
    }
  },
"TM1Data":{"Server":"Planning Sample","Cube":"plan_BudgetPlan"}})
```

You can use the CreateFromCVS method with a Common View Specification to create a Quick Report with embedded additional state information.

A Common View Specification (CVS) is a JSON that can be used to embed additional state information when creating a Quick Report. A CVS is composed of two major parts; the MDX query and a sidecar for additional state information. Data driven mechanisms, such as TurboIntegrator are only concerned with the MDX query, however user interfaces will also consume the sidecar to ensure presentation consistency. By using a CVS, you can generate highly customizable Quick Reports. For example, using a CVS, you can define aliases and subsets as per the CVS schema input.

### Syntax

The following is the syntax for the CreateFromCVS method.

`Reporting.QuickReports.CreateFromCVS(“<host system URL>”, “<server name>”, “<Common view specification>”, <boolean>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Host system URL | The host system URL where you want to generate a new report. | String
Server name | The name of the server where you want to generate a new report. | String
Common View Specification | The common view specification that you want to use to generate the new report. | String
Boolean | Set to `true` if you want the report to be generated on a new sheet at the default location. Set to `false` if you want the report to be generated in the current sheet at the default location. The `false` setting will also delete existing reports on the sheet. | True/False boolean


For more information about the Common View Specification schema, see [Commong View Specification schema](#common-view-specification-schema).


## CreateFromMDX (Quick Report)

> Example

```vb
Public Sub CreateFromMDX()
    Reporting.QuickReports.CreateFromMDX "http://vottepps06.canlab.ibm.com:9510/",
    "Planning Sample", "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].
    [Revenue]} ON 0, {[plan_time].[plan_time].[2004]} ON 1 FROM [plan_BudgetPlan]"
End Sub
```
CreateFromMDX generates a Quick Report based on the host system URL, server name, and MDX string.

### Syntax

The following string is the syntax for the CreateFromMDX method.

`Reporting.QuickReports.CreateFromMDX “<host system URL>”, “<server name>”, “<MDX>”`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Quick Report is to be created from. | Alphanumeric string
server name | Name of the server which the Quick Report is to be created from. | Alphanumeric string
MDX | MDX statement which the Quick Report is to be created from. | Alphanumeric string

## Cube

> Example

```vb
Public Sub Cube()
    MsgBox Reporting.GetCurrentReport(<ActiveCell>).Cube
End Sub
```

> If the Quick Report is located in the plan_BudgetPlan cube, in the Planning Sample server, the Cube function would return:

```vb
“{“server”:Planning Sample, “cube”:plan_BudgetPlan}”
```

Cube returns the search path of the Quick Report. 

### Syntax

The following string is the syntax for the Cube method.

`Reporting.GetCurrentReport(<ActiveCell>).Cube`

## DataSource

> Example

```vb
Public Sub DataSource()
    MsgBox Reporting.GetCurrentReport(<ActiveCell>).DataSource
End Sub
```
DataSource is used to return the Quick Report host URL.

### Syntax

The following string is the syntax for the DataSource method.

`Reporting.GetCurrentReport(<ActiveCell>).DataSource`

## EnableIndents

> Example

```vb
Public Sub EnableIndents()
    Reporting.GetCurrentReport(<ActiveCell>).EnableIndents True
End Sub
```
EnableIndents is used to enable level based indents in your Quick Report reports.

### Syntax

The following string is the syntax for the EnableIndents method.

`Reporting.GetCurrentReport(<ActiveCell>).EnableIndents <True/False value>`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
True | Enables indents in Quick Reports. | Boolean
False | Disables indents in Quick Reports. | Boolean

## ExecuteQuery

> The following syntax is an example of the ExecuteQuery method stored in a VBA module:

```vb
Public Property Get GetRowsAxis(query As String) As Collection
    Set c = Reporting.ExecuteQuery("http://pa.exampletm1.ibmcloud.com", "SData", <MDX query>)
    Dim result As New Collection
    For i = 0 To (c.GetAxes().Item(1).GetProperties().Item("tuples").GetMembers().Count - 1)
        result.Add (c.GetAxes().Item(1).GetProperties().Item("tuples").GetMembers().Item(i).GetMembers().Item(3).GetValue())
    Next i
    Set GetRowsAxis = result
End Property
```

> The following syntax is an example of the ExecuteQuery method being called in a worksheet:

```vb
Private Sub Worksheet_Change()
    Dim c As Collection
    Set c = RefreshAPIExample.GetRowsAxis(Cells(20, 4).Value2)
End Sub
```

> ExecuteQuery is triggered from a worksheet change event on cell D20. If an MDX query string exists in cell D20, and is modified, the selected MDX will be executed through the ExecuteQuery call and will return a CellSet object. This CellSet object can then be traversed in a similar way to a JSON object.

ExecuteQuery is a method used to execute selected MDX statements in your Quick Report reports.

### Syntax

The following string is the syntax for the ExecuteQuery method.

`Reporting.ExecuteQuery("<data source URL>", "<server name>", <MDX query>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
data source URL | The data source URL used in the Quick Report. | String
server name | The server name used in the Quick Report. | String
MDX query | The MDX query string to be executed by the method. | String

## GetTuple

> Example

```vb
Sub PrintTuple()
    Set tupleObject = cafe.QuickReports.Get("0").GetTuple(ActiveCell)
    Dim tuple As String
    For tupleIdx = 0 To tupleObject.Count - 1
        If tuple <> "" Then
            tuple = tuple & ", " & vbNewLine
        End If
        tuple = tuple & tupleObject.Item(tupleIdx)
    Next
    MsgBox "Tuple: " & vbNewLine & tuple
End Sub
```
GetTuple is used to return the tuple of a Quick Report at a given range. This function will return the tuple at the ActiveCell if no range is specified.

### Syntax

The following string is the syntax for the GetTuple method.

`cafe.QuickReports.Get("<Quick Report ID>").GetTuple(ActiveCell)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Quick Report ID | The ID of the Quick Report that the tuple is being returned from. | Integer


## GetSpecification

> Example

```vb
Public Sub GetSpecification()
    MsgBox Reporting.GetCurrentReport(<ActiveCell>).GetSpecification
End Sub
```
GetSpecification is used to return the MDX string that is used to build the current Quick Report.

### Syntax

The following string is the syntax for the GetSpecification method.

`Reporting.GetCurrentReport(<ActiveCell>).GetSpecification`

## GetReport

> Example

```vb
Public Sub GetReport()
    Reporting.QuickReports.Get ("5")
End Sub
```
GetReport is used to return a specific Quick Report based on the Quick Report ID.

### Syntax

The following string is the syntax for the GetReport method.

`Reporting.QuickReports.Get ("<report ID>")`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
report ID | ID of the Quick Report which the function is to return. | Integer

## ID

> Example

```vb
Public Sub ID()
    MsgBox Reporting.GetCurrentReport(<ActiveCell>).ID
End Sub
```
ID is used to return the Quick Report ID.

### Syntax

The following string is the syntax for the ID method.

`Reporting.GetCurrentReport(<ActiveCell>).ID`

## Name

> Example

```vb
Public Sub Name()
    MsgBox Reporting.GetCurrentReport(<ActiveCell>).Name
End Sub
```
Name is used to return the cube name and view name which the Quick Report is created from.

### Syntax

The following string is the syntax for the Name method.

`Reporting.GetCurrentReport(<ActiveCell>).Name`

## Rebuild

> Example

```vb
Public Sub Rebuild()
    Reporting.GetCurrentReport(<ActiveCell>).Rebuild
End Sub
```
Rebuild is used to rebuild a Quick Report.

### Syntax

The following string is the syntax for the Rebuild method.

`Reporting.GetCurrentReport(<ActiveCell>).Rebuild`

## RebuildSpecification

> Example

```vb
Public Sub RebuildSpecification()
    MsgBox Reporting.GetCurrentReport(<ActiveCell>).RebuildSpecification
End Sub
```
RebuildSpecification is used to return the MDX string that is used when rebuilding the Quick Report.

### Syntax

The following string is the syntax for the RebuildSpecification method.

`Reporting.GetCurrentReport(<ActiveCell>).RebuildSpecification`

## Refresh (Quick Report)

> Example

```vb
Public Sub Refresh()
    Reporting.GetCurrentReport(<ActiveCell>).Refresh
End Sub
```
Refresh is used to refresh a Quick Report.

### Syntax

The following string is the syntax for the Refresh method.

`Reporting.GetCurrentReport(<ActiveCell>).Refresh`

## Replace

> Example

```vb
Public Sub Replace()
    Reporting.QuickReports.Replace Reporting.GetCurrentReport(ActiveCell).4, 
    "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].[Revenue]} ON 0, 
    {[plan_time].[plan_time].[2004]} ON 1 FROM [plan_BudgetPlan]"
End Sub
```
Replace is used to replace the MDX statement in the Quick Report with another MDX statement.

### Syntax

The following string is the syntax for the Replace method.

`Reporting.QuickReports.Replace Reporting.GetCurrentReport(ActiveCell).<Quick Report ID>, <MDX statement>`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Quick Report ID | The ID of the Quick Report that will have its MDX statement replaced. | Integer
MDX statement | The MDX statement that will be replacing the current MDX statement in the Quick Report. | String

## ReplaceWithFormats

> Example

```vb
Public Sub ReplaceWithFormats()
   Reporting.QuickReports.ReplaceWithFormats Reporting.GetCurrentReport(ActiveCell).4,
   "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].[Revenue]} ON 0,
   {[plan_time].[plan_time].[2004]} ON 1 FROM [plan_BudgetPlan]", True
End Sub
```
ReplaceWithFormats is used to replace the MDX statement in the Quick Report with another MDX statement. ReplaceWithFormats also has the option to preserve or destroy the existing sheet formatting in the Quick Report.

### Syntax

The following string is the syntax for the ReplaceWithFormats method.

`Reporting.QuickReports.ReplaceWithFormats Reporting.GetCurrentReport(ActiveCell).<Quick Report ID>, <MDX statement>, <ReFormat>`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Quick Report ID | The ID of the Quick Report that will have its MDX statement replaced. | Integer
MDX statement | The MDX statement that will be replacing the current MDX statement in the Quick Report. | String
ReFormat | Defines whether or not to preserve or destroy the sheet formatting in the existing Quick Report. `True` preserves the sheet formatting. `False` destroys the sheet formatting. | Boolean

## RowHierarchies

> Example

```vb
Sub RowHierarchies()
    Dim slicers As String
    For Each Slicer In cafe.QuickReports.Get("0").SlicerDimensions
        If slicers <> "" Then
            slicers = slicers & ", " & vbNewLine
        End If
        slicers = slicers & Slicer
    Next
    MsgBox "Rows:" rows 
End Sub
```
RowHierarchies is used to return the hierarchies that exist in the rows of a Quick Report.

### Syntax

The following string is the syntax for the RowHierarchies method.

`cafe.QuickReports.Get("<Quick Report ID").RowDimensions`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Quick Report ID | The ID of the Quick Report that the row hierarchies are being returned from. | Integer

## Select

> Example

```vb
Public Sub SelectReport()
    Reporting.GetCurrentReport(<ActiveCell>).Select
End Sub
```
Select is used to select and highlight the current active Quick Report.

### Syntax

The following string is the syntax for the Select method.

`Reporting.GetCurrentReport(<ActiveCell>).Select`

## SetSlicer

> Example

```vb
Public Sub SetSlicer()
    Reporting.GetCurrentReport(<ActiveCell>).SetSlicer "[plan_business_unit].
    [plan_business_unit]", "10100"
End Sub
```
SetSlicer is used to set the values for a slicer dimension in the Quick Report.

### Syntax

The following string is the syntax for the SetSlicer method.

`Reporting.GetCurrentReport(<ActiveCell>).SetSlicer “<dimensions>, <name>”`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
dimensions | The dimensions to set the slicer to. | String
name | The name to set the slicer to. | String

## SlicerHierarchies

> Example

```vb
Sub RowHierarchies()
    Dim slicers As String
    For Each Slicer In cafe.QuickReports.Get("0").SlicerDimensions
        If slicers <> "" Then
            slicers = slicers & ", " & vbNewLine
        End If
        slicers = slicers & Slicer
    Next
    MsgBox "Slicers:" slicers
End Sub
```
SlicerHierarchies is used to return the hierarchies that exist in the slicers of a Quick Report.

### Syntax

The following string is the syntax for the SlicerHierarchies method.

`cafe.QuickReports.Get("<Quick Report ID").SlicerDimensions`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Quick Report ID | The ID of the Quick Report that the slicer hierarchies are being returned from. | Integer

## UseServerFormats

> Example

```vb
Public Sub ToggleServerFormats(r As Range)
   r.Worksheet.Activate
   Set fView = Reporting.GetCurrentReport(r)
   If Not (fView Is Nothing) Then
     fView.UseServerFormats = Not fView.UseServerFormats
     fView.Refresh
   End If
End Sub
```
UseServerFormats clears any user applied formatting and applies server based formatting after a Quick Report is refreshed.

### Syntax

The following string is the syntax for the UseServerFormats method.

`Reporting.GetCurrentReport(<ActiveCell>).UseServerFormats = <True/False>`
