# Exploration API functions

Exploration functions can be used to interact with exploration worksheets. 

Exploration functions can use the following PropertyAccessor objects:

PropertyAccessor | Description | 
--------- | ------- | -----------
Count | Counts the number of Explorations in active book.
GetAt(sheet) | Gets the Exploration object on the specified sheet name, from the active book, if it exists.
GetReports() | Gets the collection of Exploration objects from the active book.

The Exploration functions that are exposed through the IBM® Cognos® automation objects are:
 
## Clear (Exploration)

> Example

```vb
Public Sub Clear()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).Clear
End Sub
```
Clear is used to clear all of the data values in the exploration.

### Syntax

The following string is the syntax for the Clear method.

`Reporting.Explorations.GetAt().Clear`

## Create (Exploration)

> Example

```vb
Public Sub Create()
    Reporting.Explorations.create "http://computername", "Planning Sample", 
    "plan_BudgetPlan", "Goal Input"
End Sub
```
Create generates an Exploration View based on the host system URL, server name, cube name, and view name.

### Syntax

The following string is the syntax for the Create method.

`Explorations.Create "<host system URL>", "<server name>", "<cube name>", "<view name>" `

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Exploration View is to be created from. | Alphanumeric string
server name | Name of the server which the Exploration View is to be created from. | Alphanumeric string
cube name | Name of the cube which the Exploration View is to be created from. | Alphanumeric string
view name | Name of the view which the Exploration View is to be created from. | Alphanumeric string

## CreateFromMDX (Exploration)

> Example

```vb
Public Sub CreateFromMDX()
    Reporting.Explorations.CreateFromMDX "http://vottepps06.canlab.ibm.com:9510/", 
    "Planning Sample", "SELECT {[plan_chart_of_accounts].[plan_chart_of_accounts].
    [Revenue]} ON 0, {[plan_time].[plan_time].[2004]} ON 1 FROM [plan_BudgetPlan]"
End Sub
```
CreateFromMDX generates an Exploration View based on the host system URL, server name, and MDX string.

### Syntax

The following string is the syntax for the CreateFromMDX method.

`Reporting.Explorations.CreateFromMDX “<host system URL>”, “<server name>”, “<MDX>”`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
host system URL | URL of the host system which the Exploration View is to be created from. | Alphanumeric string
server name | Name of the server which the Exploration View is to be created from. | Alphanumeric string
MDX | MDX statement which the Exploration View is to be created from. | Alphanumeric string

## GetColumnSuppression

> Example

```vb
Public Sub AreColumnsSuppressed()
    MsgBox Reporting.Explorations.GetAt(Application.ActiveSheet.Name).
    GetColumnSuppression
End Sub
```
GetColumnSuppression is used to return whether or not zero-suppression is applied to columns in the exploration.

### Syntax

The following string is the syntax for the GetColumnSuppression method.

`Reporting.Explorations.GetAt().GetColumnSuppression`

## GetRowSuppression

> Example

```vb
Public Sub AreRowsSuppressed()
    MsgBox Reporting.Explorations.GetAt(Application.ActiveSheet.Name).
    GetRowSuppression
End Sub
```
GetRowSuppression is used to return whether or not zero-suppression is applied to rows in the exploration.

### Syntax

The following string is the syntax for the GetRowSuppression method.

`Reporting.Explorations.GetAt().GetRowSuppression`

## GetSpecification

> Example

```vb
Public Sub GetSpecification()
    Msgbox
Reporting.Explorations.GetAt(Application.ActiveSheet.Name).GetSpecification
End Sub
```
GetSpecification is used to return the MDX string that is used to build the current Exploration.

### Syntax

The following string is the syntax for the GetSpecification method.

`Reporting.Explorations.GetAt().GetSpecification`

## GetValue

> Example

```vb
Public Sub ToggleSetEditorPreview()
    Dim x
    x = Reporting.Settings.GetValue("SetEditorPreviewOn")
    If "True" = x Then
        Reporting.Settings.SetValue "SetEditorPreviewOn", "False"
    Else
        Reporting.Settings.SetValue "SetEditorPreviewOn", "True"
    End If  
End Sub
```
GetValue is used to retrieve the value of a particular setting in a session.

### Syntax

The following string is the syntax for the GetValue method.

`Reporting.Settings.GetValue("<Setting>")`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Setting | The name of the setting whose value you want to retrieve. | String

## Refresh (Exploration)

> Example

```vb
Public Sub Refresh()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).Refresh
End Sub
```
Refresh is used to refresh the exploration.

### Syntax

The following string is the syntax for the Refresh method.

`Reporting.Explorations.GetAt().Refresh`

## SwapRowsAndColumns

> Example

```vb
Public Sub SwapsRowsAndColumns()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).SwapsRowsAndColumns
End Sub
```
SwapRowsAndColumns is used to swap the rows and columns in an exploration.

### Syntax

The following string is the syntax for the SwapRowsAndColumns method.

`Reporting.Explorations.GetAt().SwapRowsAndColumns`

## SetRowSuppression

> Example

```vb
Public Sub SetRowSuppressions()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).SetRowSuppression 
    True
End Sub
```
SetRowSuppression is used to enable and disable zero-suppression for rows in an exploration.

### Syntax

The following string is the syntax for the SetRowSuppression method.

`Reporting.Explorations.GetAt().SetRowSuppression <True/False value>`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
True | Enables zero-suppression. | Boolean
False | Disables zero-suppression. | Boolean

## SetColumnSuppression

> Example

```vb
Public Sub SetColumnSuppressions()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).SetColumnSuppression 
    True
End Sub
```
SetColumnSuppression is used to enable and disable zero-suppression for columns in an exploration.

### Syntax

The following string is the syntax for the SetColumnSuppression method.

`Reporting.Explorations.GetAt().SetColumnSuppression <True/False value>`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
True | Enables zero-suppression. | Boolean
False | Disables zero-suppression. | Boolean

## SetSpecification

> Example

```vb
Public Sub SetSpecifications()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).SetSpecification 
    "SELECT TM1SubsetToSet([plan_time], ""current_year_and_qtrs"") DIMENSION 
    PROPERTIES MEMBER_UNIQUE_NAME, MEMBER_NAME, MEMBER_CAPTION, LEVEL_NUMBER, 
    CHILDREN_CARDINALITY, [plan_time].[Time] ON 0, TM1TOGGLEDRILLSTATE
    (TM1SubsetToSet([plan_chart_of_accounts], ""Default"") , 
    {[plan_chart_of_accounts].[Revenue],[plan_chart_of_accounts].
    [Operating Expense]} , EXPAND_BELOW , RECURSIVE) DIMENSION PROPERTIES 
    MEMBER_UNIQUE_NAME, MEMBER_NAME, MEMBER_CAPTION, LEVEL_NUMBER, 
    CHILDREN_CARDINALITY, [plan_chart_of_accounts].[AccountName] ON 1 FROM 
    [plan_BudgetPlan] WHERE ([plan_version].[FY 2004 Budget] , 
    [plan_business_unit].[10000] , [plan_department].[1000] , 
    [plan_exchange_rates].[actual] , [plan_source].[goal]) DIMENSION PROPERTIES 
    MEMBER_UNIQUE_NAME, MEMBER_NAME, MEMBER_CAPTION, LEVEL_NUMBER, 
    CHILDREN_CARDINALITY , [plan_version].[VersionName] , [plan_business_unit].
    [BusinessUnit] , [plan_department].[Department] , [plan_source].[Source]"
End Sub
```
SetSpecification is used to define the subset and dimension properties of an existing exploration.

### Syntax

The following string is the syntax for the SetSpecification method.

`Reporting.Explorations.GetAt().SetSpecification “<MDX>”`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
MDX | MDX statement used to define the subset and dimension properties of the exploration. | String

## SetValue

> Example

```vb
Public Sub ToggleSetEditorPreview()
    Dim x
    x = Reporting.Settings.GetValue("SetEditorPreviewOn")
    If "True" = x Then
        Reporting.Settings.SetValue "SetEditorPreviewOn", "False"
    Else
        Reporting.Settings.SetValue "SetEditorPreviewOn", "True"
    End If  
End Sub
```
SetValue is used to set a new value for a specific setting and save the changes to the settings file.

### Syntax

The following string is the syntax for the SetValue method.

`Reporting.Settings.SetValue "<Setting>", "<Value>"`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
Setting | The name of the setting whose value you want to set. | String
Value | The boolean value you want to set for the specified setting. | True/False boolean


## Unlink

> Example

```vb
Public Sub Unlink()
    Reporting.Explorations.GetAt(Application.ActiveSheet.Name).Unlink
End Sub
```
Unlink is used to convert an exploration to a static worksheet.

### Syntax

The following string is the syntax for the Unlink method.

`Reporting.Explorations.GetAt().Unlink`