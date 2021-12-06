# Global API functions

 Global API functions can be used to interact with any IBM Planning Analytics for Microsoft Excel worksheets. The global functions that are exposed through the IBM Cognos® automation objects are:

## GetConnection

GetConnection is a method exposed by the top level reporting API object. If you want to use the REST APIs, you'll need to use the GetConnection method to return the connection object that implements the REST request methods (GET, POST, DELETE, PATCH) and communicate with the TM1 Server.

To learn more about REST request methods, see [REST API](#rest-api).

### Syntax

The following string is the syntax for the GetConnection method. 
To use the method, you must know the URL of the host that you want to send REST requests to.
`Reporting.GetConnection(<CURRENT>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
CURRENT | The URL of the host that you want to send REST requests to. | String

## ChangeDataSource (Task Pane)

You can use the ChangeDataSource method to change datasources within a session. You might be prompted for a login if you weren't logged in. 

### Syntax

Reporting.TaskPane.ChangeDatasource "host system URL" "server name"


## ClearAllData

> Example

```vb
CognosOfficeAutomationObject.ClearAllData 
```

ClearAllData clears all data values in the opened workbooks.

### Syntax

The following string is the syntax for the ClearAllData method.

`ClearAllData()`

## ClearBook

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").ClearBook
```

ClearBook clears Planning Analytics for Microsoft Excel data in the active book.

### Syntax

The following string is the syntax for the ClearBook method.

`ClearBook()`

## ClearCache

> Example

```vb
CognosOfficeAutomationObject.ClearCache()
```

ClearCache reduces the size of an IBM® Planning Analytics for Microsoft Excel workbook by clearing metadata and data from formulas.

### Syntax

The following string is the syntax for the ClearCache method.

`ClearCache()`

## ClearSelection

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").ClearSelection
```

ClearSelection clears IBM Planning Analytics for Microsoft Excel data in the active selection.

### Syntax

The following string is the syntax for the ClearSelection method.

`ClearSelection()`

## ClearSheet

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").ClearSheet
```

ClearSheet clears IBM Planning Analytics for Microsoft Excel data in the active sheet.

### Syntax

The following string is the syntax for the ClearSheet method.

`ClearSheet()`

## EvaluateSynchronous

The EvaluateSynchronous function behaves like the native Excel Evaluate formula and turns a string into a formula. For example, you can use EvaluateSynchronous to find the nth child of a parent or to determine the parent of a given dimension element.

### Syntax

The following string is the syntax for the EvaluateSynchronous method.

`Reporting.EvaluateSynchronous(evalString)`

### SinglePassMode 

To use Excel’s Evaluate formula instead of EvaluateSynchronous, you can toggle single pass mode on and off between evaluations with Reporting.SinglePassMode. 

You must set the SinglePassMode flag to false before the Exit Function to avoid any performance degradation.

> Example for turning SinglePassMode on and off:

```vb
Reporting.SinglePassMode = true
var X = Evaluate("=ELCOMPN(""" & cube & ":" & dimension & """," & """" & parent & """" & ")")
Reporting.SinglePassMode = false  
Exit Function
 
```
#### Syntax

`Reporting.SinglePassMode`

## Hide (Task Pane)

> Example:

```vb
Dim bResult As Boolean
bResult = Reporting.TaskPane.IsVisible()
If bResult = True Then
Call Reporting.TaskPane.Hide()
End If
```

You can use the Hide method to hide the Task Pane. 

### Syntax

The following is the syntax for the Hide method.

`Reporting.TaskPane.Hide()`

## HttpLogonCredentials

The HttpLogonCredentials function authenticates a user to a Web site that requires new authentication credentials, such as Basic, Kerberos, and SiteMinder. HttpLogonCredentials takes the URL, user name, and password that are used for authentication on the Web site.

### Syntax

IBM® Cognos® does not support SiteMinder form-based authentication. You must use the IBM Cognos menu commands and options instead of the API to automate the refreshing and publishing of content.

`HttpLogonCredentials (url, user name, password)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
url | The URL for the Web site against which you want to authenticate. | String
user name | The user name for authentication. | String
password | The password for authentication. | String

## IsVisible (Task Pane)

> Example:

```vb
Dim bResult As Boolean
bResult = Reporting.TaskPane.IsVisible()
If bResult = True Then
Call Reporting.TaskPane.Refresh()
End If
```

You can use the IsVisible method to return the state of the Task Pane. If True is returned, the Task Pane is visible. If False is returned, the Task Pane is not visible.

### Syntax

The following is the syntax for the IsVisible method.

`Reporting.TaskPane.IsVisible()`

## Logoff

> Example

```vb
CognosOfficeAutomationObject.Logoff
```

Logoff logs off all the IBM® Cognos® servers to which users are currently logged on.

### Syntax

The following string is the syntax for the Logoff method.

`Logoff()`

## Logon

> Example of the syntax for logging into an IBM Cognos Analytics system: 

```vb
Dim bResult As Boolean

bResult = CognosOfficeAutomationObject.Logon
("http://localhost/ibmcognos/cgi-bin/cognos.cgi",
"Administrator", "CognosAdmin", "Production")
```

> Example of the syntax for logging into an IBM TM1 system: 

```vb
Dim bResult As Boolean
bResult = CognosOfficeAutomationObject.Logon
("http://myPlanningAnalyticsServer.com",
"admin", "peaches", "localhost/Planning Sample")
```

The Logon function takes the URL of the server and the credential elements required by IBM® Planning Analytics for Microsoft Excel to perform a logon: user ID, password, and namespace. The namespace parameter is case-sensitive; therefore, you must match the namespace exactly. Planning Analytics for Microsoft Excel uses the Logon function, whether you're logging into an IBM Cognos Analytics system or an IBM TM1 system.

IBM Cognos® Office stores user credentials only in memory. For this reason, users are responsible for storing their credentials in a secured area and passing them to the logon methods at run time.

If you use the Logon function with incorrect credentials, the system raises a CAMException error, however, no exception is written to the log file indicating a failure. To avoid this situation, remember that strings are case-sensitive and ensure that you use valid user IDs, passwords, and namespaces.

Logon does not appear in the macro list in the Microsoft application because the macro receives an argument. Any macro with parameters is by definition private and private macros are not shown in the macro options by default.

<aside class="notice">
The Logon function cannot be used to log into a cloud-based system.
</aside>

### Syntax

The following string is the syntax for the Logon method.

`Boolean Logon (url, user name, password, namespace)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
url | The URL for the IBM Cognos Analytics or IBM TM1 system, which you want to log on to. | String
user name | The user name for authentication. | String
password | The password for authentication. | String
namespace | The specific namespace for authentication. | String

<aside class="notice">
Mode 1 authentication requires the combination of the PM Hub host and the TM1 Server as the namespace, separated by a forward slash (/). If your namespace contains a forward slash, the logon is interpreted as a Mode 1 authentication attempt.
</aside>

### Return value

Data type: Boolean

The Boolean value is true if successful

## Publish

> Example of the syntax for publishing to a IBM Cognos Analytics data source:

```vb
Publish("CAMID('::Anonymous')/folder[@name='My
Folders']","Description of 'My Folders'", "")
```

> Example of the syntax for publishing to a IBM Planning Analytics data source:

```vb
("https://myPAconnection.PlanningAnalytics.com", "C:\path\to\local\file.xlsx", "/tm1/Planning%20Sample/api/v1/Contents('Applications')/Contents('Planning %20Sample')", "PublishedFileName.xlsx", "My Description", "MyToolTip")
```

Use Publish to publish content to IBM® Cognos® Connection or to a TM1 Server Application Folder. 

### Syntax

The arguments mirror the entry boxes in the dialog box that is used in the user interface.

Publish does not appear in the macro list in the Microsoft application because the macro receives an argument. Any macro with parameters is by definition private and private macros are not shown in the macro options by default.

`Publish (url, document path, server path, name, description, screenTip)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
url | The server to which you are publishing. | String
document path | The location of the document to be published. It is the local path of the file that you want to publish. If the path of your folder is not correct when you publish using automation, you are again prompted to log on. This is because IBM Cognos does not distinguish between non-existing folders and folders for which the user does not have permissions. This security feature helps to prevent the discovery of the folder path by trial and error.

In IBM Cognos Analytics, the folder path is a search path. For more information, see the IBM Cognos Analytics Administration Guide. | String
server path | The path in the content store where the document is saved. | String
name | The document name that will appear in IBM Cognos. | String
description | The document description that will appear in IBM Cognos. | String
screenTip | The text that users see when they point to the document in IBM Cognos. | String

## PublishTm1

> Example for publishing to a TM1 data source:

```vb
Public Sub PublishTm1()
    Dim oMessageSuppressor As CognosOfficeMessageSuppressor
    Set oMessageSuppressor = New CognosOfficeMessageSuppressor
    Dim sUrl As String
    sUrl = "http://myserver.ibm.com"
    Dim sDS As String
    sDS = "Planning Sample"
    Dim sPublishPath As String
    sPublishPath = "Contents('Planning Sample')/Contents('Top Down Goals')"
    Dim sDocumentPath As String
    sDocumentPath = "C:\Users\JC\Documents\Publish API book.xlsx"
    Dim sName As String
    sName = "Publish&Tm1 test1"
    Dim sDescription As String
    sDescription = "New PublishTm1 api"
    Dim sScreenTip As String
    sScreenTip = "test"
    Dim sIsPrivate As Boolean
    sIsPrivate = True
    Dim sAsReference As Boolean
    sAsReference = False
  
    CognosOfficeAutomationObject.PublishTm1 sUrl, sDS, sPublishPath, sDocumentPath, sName, sDescription, sScreenTip, sIsPrivate, sAsReference
    Exit Sub
End Sub 
```

PublishTm1 is a TM1-specific API that differs from the existing _Publish_ api in the following ways:

+ No need to include `/tm1/Planning%20Sample/api/v1/Contents('Applications')/` in the publish path; the API fills that in during execution.
+ No need to encode spaces and other special characters.
+ Takes a Boolean argument to control publish scope (public/private).
+ Takes a Boolean argument to publish as reference.

### Syntax

`PublishTm1(string serverURL, string serverName, string publishPath, string documentPath, string name, string description, string screenTip, bool isPrivate, bool asReference)`

## MakeFolderTm1

> Example for creating a public or private folder at a given location:

```vb
Public Sub MakeFolderTm1()
    Dim oMessageSuppressor As CognosOfficeMessageSuppressor 
    Set oMessageSuppressor = New CognosOfficeMessageSuppressor
    Dim sUrl As String
    sUrl = "http://ablauts1.fyre.ibm.com"
    Dim sDS As String
    sDS = "Planning Sample"
    Dim sFolderPath As String
    sFolderPath = "Contents('Planning Sample')/Contents('Top Down Goals')"
    Dim sFolderName As String
    sFolderName = "PublishTm1_Folder1"
    Dim sIsPrivate As Boolean
    sIsPrivate = False
  
    CognosOfficeAutomationObject.MakeFolderTm1 sUrl, sDS, sFolderPath, sFolderName, sIsPrivate
    Exit Sub
End Sub 
```

MakeFolderTm1 is a TM1-specific API that can create a public or private folder at a given location and differs from the existing _Publish_ api in the following ways:

+ No need to include `/tm1/Planning%20Sample/api/v1/Contents('Applications')/` in the publish path; the API fills that in during execution.
+ No need to encode spaces and other special characters.
+ Takes a Boolean argument to control publish scope (public/private).

### Syntax

`MakeFolderTm1(string serverURL, string serverName, string folderPath, string folderName, bool isPrivate)`

## PublishTm1Multiple

> Example for publishing multiple files to a TM1 data source:

```vb
Public Sub PublishTm1Multiple()
On Error GoTo HANDLER
    Dim oMessageSuppressor As CognosOfficeMessageSuppressor
        'Use the message suppressor to turn off all Cognos Office messages.
    Set oMessageSuppressor = New CognosOfficeMessageSuppressor
    Dim sUrl As String
    sUrl = "http://ablauts1.fyre.ibm.com"
    Dim sDS As String
    sDS = "Planning Sample"
    Dim sPublishPath As String
    sPublishPath = "Contents('Planning Sample')/Contents('Top Down Goals')"
    Dim sDocumentPaths() As String
    sDocumentPaths = Split("C:\Users\JC\Documents\Publish API testbook.xlsx, C:\Users\JC\Documents\4420.xlsx", ",")
    Dim sNames() As String
    sNames = Split("Publish Multiple test 1,Publish Multiple test 2", ",")
    Dim sDescriptions() As String
    sDescriptions = Split("A test of the new PublishTm1 api", ",")
    Dim sScreenTips As String
    sScreenTips() = Split("A test", ",")
    Dim sIsPrivate As Boolean
        'To publish with a private scope set this to true
    sIsPrivate = True
    Dim sAsReference As Boolean
        'To publish the file as a reference set this to true
    sAsReference = False
  
        'Call the Cognos Office Automation object to publish the file.  
    CognosOfficeAutomationObject.PublishTm1Multiple sUrl, sDS, sPublishPath, sDocumentPaths, sNames, sDescriptions, sScreenTips, sIsPrivate, sAsReference
    Exit Sub
    HANDLER:  
End Sub 
```

PublishTm1Multiple is a TM1-specific API that is an expansion on the _PublishTm1_ api, which allows a list of files to be published at once to the same location with the same scope. _PublishTm1Multiple_ differs from the existing _Publish_ api in the following ways:

+ No need to include `/tm1/Planning%20Sample/api/v1/Contents('Applications')/` in the publish path; the API fills that in during execution.
+ No need to encode spaces and other special characters.
+ Takes a string array of file locations and names (must be a 1-to-1 pairing of files and file names).
+ Takes a string array of descriptions and screentips. If these aren’t provided API defaults to an empty string.
+ Takes a Boolean argument to control publish scope (public/private).
+ Takes a Boolean argument to publish as reference.
+ Publish location, scope, and as reference settings are applied to all the files in the array.

### Syntax

`PublishTm1Multiple(string serverURL, string serverName, string publishPath, string[] documentPaths, string[] names, string[] descriptions, string[] screenTips, bool isPrivate, bool asReference)`

## Refresh (Task Pane)

> Example:

```vb
Dim bResult As Boolean
bResult = Reporting.TaskPane.IsVisible()
If bResult = True Then
Call Reporting.TaskPane.Refresh()
End If
```

You can use the Refresh method to refresh the metadata tree in the Task Pane. 

### Syntax

The following is the syntax for the Refresh method.

`Reporting.TaskPane.Refresh()`

## RefreshAllData

> Example

```vb
Dim bResult as Boolean
Copy

bResult = CognosOfficeAutomationObject.Logon
("http://localhost/ibmcognos/cgi-bin/cognos.cgi",
"Administrator", "CognosAdmin", "Production")
Copy

'Refresh the data if we successfully logged on to the
IBM Cognos server.
Copy

If bResult Then
Copy

  CognosOfficeAutomationObject.RefreshAllData
Copy

End If
```

 RefreshAllData fetches the most current data values from the IBM® TM1 server and updates those values in the current document. 

### Syntax

The system must be successfully logged on to the IBM TM1 server.

If you are using IBM Cognos Office with IBM Cognos® Analytics data, ensure that the Prompt Update Method property on the Manage Data tab in the IBM Cognos pane is set to Use=Display or Do Not Update to complete the operation. Otherwise, the report cannot be refreshed without user intervention and generates errors.

`RefreshAllData()`

## RefreshAllDataAndFormat

> Example

```vb
Dim bResult as Boolean
Copy

bResult = CognosOfficeAutomationObject.Logon
("http://localhost/ibmcognos/cgi-bin/cognos.cgi",
"Administrator", "CognosAdmin", "Production")
Copy

'Refresh the data and formatting if we successfully logged on to the
IBM Cognos server.
Copy

If bResult Then
Copy

  CognosOfficeAutomationObject.RefreshAllDataAndFormat
Copy

End If
```

RefreshAllDataAndFormat retrieves the most current data values and formatting from the IBM® Cognos® server and updates those values and formats in the current document.

### Syntax

The system must be successfully logged on to the IBM Cognos server.

If you are using IBM Cognos Office with IBM Cognos Analytics data, ensure that the **Prompt Update Method** property on the **Manage Data** tab in the IBM Cognos pane is set to **Use=Display** or **Do Not Update** to complete the operation. Otherwise, the report cannot be refreshed without user intervention and generates errors.

`RefreshAllDataAndFormat()`

## RefreshBook

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").RefreshBook
```

RefreshBook refreshes all data values in the opened workbooks.

### Syntax

The following string is the syntax for the RefreshBook method.

`RefreshBook()`

## RefreshSelection

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").RefreshSelection
```

 RefreshSelection refreshes IBM Planning Analytics for Microsoft Excel data in the active selection. 
 
### Syntax

The following string is the syntax for the RefreshSelection method.

`RefreshSelection()`


## RefreshSheet

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").RefreshSheet
```

 RefreshSheet refreshes IBM Planning Analytics for Microsoft Excel data in the active sheet. 

### Syntax

The following string is the syntax for the RefreshSheet method.

`RefreshSheet()`

## Settings

> Example using `SetValue`

```vb
Reporting.Settings.SetValue "GroupingOption", "Full"
```

> Example using `GetValue`

```vb
Reporting.Settings.GetValue ("ShowServerInExploration")
```

The Settings function can be used to enable, disable, or define settings in Planning Analytics for Microsoft Excel. 

### Syntax

SetValue is used to set a value in a setting.

`Reporting.Settings.SetValue "<setting name>", "<setting value>"`

GetValue is used to retrieve a value of a setting.

`Reporting.Settings.GetValue("<setting name>")`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
setting name | The name of the setting that you want to enable, disable, or define. | Alphabetic
setting value | The value that you want to use to enable, disable, or define in the setting. | Alphabetic, alphanumeric, boolean, integer

View Settings in the [CognosOfficeReportingSettings.xml](https://www.ibm.com/support/knowledgecenter/SSD29G_2.0.0/com.ibm.swg.ba.cognos.ug_cxr.2.0.0.doc/c_cognosofficereportingsettings.html) file for a list of the possible settings and values that you can use.

## Show (Task Pane)

> Example:

```vb
Dim bResult As Boolean
bResult = Reporting.TaskPane.IsVisible()
If bResult = True Then
Call Reporting.TaskPane.Show()
End If
```

You can use the Show method to reveal the Task Pane. 

### Syntax

The following is the syntax for the Show method.

`Reporting.TaskPane.Show()`

## SuppressMessages

> Example

```vb
Private Sub Class_Initialize()
    CognosOfficeAutomationObject.SuppressMessages True
End Sub
Private Sub Class_Terminate()
    CognosOfficeAutomationObject.SuppressMessages False
End Sub
```

When added to an existing script or function in the Planning Analytics for Microsoft Excel API, SuppressMessages suppresses all of the messages and dialog boxes that may arise from the script or function.

In addition to SuppressMessage, you need to set the Application.DisplayAlerts property in Microsoft Excel to `false`. For more information about the Application.DisplayAlerts property, see [Application.DisplayAlerts property](https://docs.microsoft.com/en-us/office/vba/api/excel.application.displayalerts).

### Syntax

The following string is the syntax for the SuppressMessages method.

`SuppressMessages()`

## TraceError

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.TraceError("VBA method failed")
```
> The following is an example of the appended error information in the IBM Planning Analytics for Microsoft Excel log file:

```vb
[Severity=Error]
[Exception] TraceError(String error)
[Thread=6, Background=True, Pool=True, Domain=]
[System.Exception] VBA API ERROR: VBA method failed
```

TraceError appends error information into the IBM Planning Analytics for Microsoft Excel log file. The user defines the error information they wish to append to the log file for errors.

### Syntax

The following string is the syntax for the TraceError method.

`TraceError("<user defined error information>")`

## TraceLog 

> Example

```vb
Dim strTraceLog as String
strTraceLog = CognosOfficeAutomationObject.TraceLog
MsgBox strTraceLog
```

TraceLog returns all the automation activities and errors.

### Syntax

The following string is the syntax for the TraceLog  method.

`*String* TraceLog ()`

### Return Value

Data type: String

The value of the logging item as string

## UnlinkAllData

> Example

```vb
CognosOfficeAutomationObject.UnlinkAllData
```

UnlinkAllData disconnects all the IBM® Cognos® data values in the current document. The values are no longer updated with subsequent calls to RefreshAllData. The values become static. 

### Syntax

For IBM Cognos Office, any IBM Cognos data values that are imported into the current document after UnlinkAllData is called will continue to be linked to the IBM Cognos data source.

The values can be updated with new server data using the RefreshAllData call.

`UnlinkAllData ()`

## UnlinkBook 

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").UnlinkBook 
```

UnlinkBook unlinks the active book from the connection.

### Syntax

The following string is the syntax for the UnlinkBook method.

`UnlinkBook()`

## UnlinkSelection

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").UnlinkSelection
```

UnlinkSelection disconnects the selected data values. The values are no longer updated with subsequent calls to Refreshable. The values become static.

### Syntax

The following string is the syntax for the UnlinkSelection method.

`UnlinkSelection()`

## UnlinkSheet 

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").UnlinkSheet 
```

UnlinkSheet unlinks the active sheet from the connection.

### Syntax

The following string is the syntax for the UnlinkSheet  method.

`UnlinkSheet()`

## UpdateServerUrl

> Example

```vb
UpdateServerUrl "http://testserver1/cgi-bin/cognos.cgi" 
"http://prodserver1/cgi-bin/cognos.cgi"
```

> The following example uses only the part of the URL that is changing:

```vb
UpdateServerUrl "testserver1" "prodserver1"
```

Use UpdateServerUrl to update the IBM® Cognos® server information for existing reports and formulas. 

### Syntax

The UpdateServerUrl method takes two arguments: the old server URL and the new server URL. These arguments mirror the entry boxes in the **Update System** dialog box. To gain access to this control from IBM Cognos, click the **Options** button on the IBM Cognos ribbon, then click **Update System Utility**.

The UpdateServerUrl method replaces the server information for existing reports. When running this command, the name of the package or data source remains the same. You can use this method to change only one server, such as a test server to a production server. The URL arguments can be full or partial URLs. If any argument is empty, this command does nothing, however, running this command with empty arguments has the potential to corrupt the report. Server information is stored in both the server property and the serialized report property. Running an empty command could cause these two instances to get out of sync.

Because the UpdateServerUrl method searches and replaces strings, it is possible to use only part of the URL, provided it is a unique substring. 

<aside class="notice">
The UpdateServerUrl search looks at all data in the workbook and updates data that matches the search string, not just report properties containing the URL string. Therefore, when you use only part or all of the original URL string with the UpdateServerUrl method, you will change all data that matches the search string.
</aside>

`UpdateServerUrl "old server URL string" "new server URL string"`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
old server URL string | Indicates the URL of the source or current system. | String
new server URL string | Indicates the URL of the target system. | String

## UserAgent

> Example

```vb
  Dim useragent as String
  useragent = Reporting.UserAgent
```

Returns the product and build version details. For example, `PAfE/2.0.66.9 (8590999552); Excel/16.0.13127`.

### Syntax
The following string is the syntax for the UserAgent method. 
`Reporting.UserAgent`

## UserAgentSCRelease

> Example

```vb
  Dim release as String
  release = Reporting.UserAgentSCRelease
```

Returns the condensed product version details. For example, `66.9`.

### Syntax
The following string is the syntax for the UserAgentSCRelease method.
`Reporting.UserAgentSCRelease`

## UserAgentSCReleaseFull

> Example

```vb
  Dim releasefull as String
  releasefull = Reporting.UserAgentSCReleaseFull
```

Returns the full product version details. For example, `2.0.66.9`.

### Syntax
The following string is the syntax for the UserAgentSCReleaseFull method.
`Reporting.UserAgentSCReleaseFull`


## Wait

> Example

```vb
Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").Wait
```

> Usage example

```vb
Sub Wait()
    Reporting.GetCurrentReport(ActiveCell).Commit
    Reporting.Wait
    Reporting.GetCurrentReport(ActiveCell).Refresh
End Sub
Sub Wait()
    Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").RefreshBook
    Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer.Application("COR", "1.1").Wait
    MsgBox "Refresh complete!"
End Sub
```

Wait holds the VBA thread until all prior IBM Planning Analytics for Microsoft Excel background tasks are complete. 

### Syntax

The following string is the syntax for the Wait method.

`Wait()`
