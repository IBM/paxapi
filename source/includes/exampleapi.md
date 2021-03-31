#Examples of processing

## Processing outside of VBA

> The following Visual Basic Script opens Microsoft Office Excel, logs on to IBM Cognos Analytics, refreshes the content, and logs off.

```vb
' Start Excel in batch mode

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = False

objExcel.ScreenUpdating = False

objExcel.DisplayAlerts = False

'Open a workbook that has IBM Cognos data
in it.

Set objWorkbook = objExcel.Workbooks.Open("C:\workbook1.xls")

' Call the wrapper macros

bResult = CognosOfficeAutomationObject.Logon
("http://localhost/ibmcognos/cgi-bin/cognos.cgi",
"Administrator", "CognosAdmin", "Production")

objExcel.Run "RefreshAllData"

objExcel.Run "Logoff"

objExcel.Run "TraceLog", "C:\AutomationLog.log"

objWorkbook.Save

objWorkbook.Close

objExcel.Quit
```

 If you want to use IBM® Cognos® Office Automation outside VBA, you cannot call the APIs directly. This topic describes how you can call the APIs outside of VBA.

To use IBM Cognos Office Automation outside VBA, you must create wrapper macros in the Microsoft Office document for each API. You can then call these macros from your code. The module CognosOfficeAutomationExample.bas is an example of a wrapper macro that you can call from outside VBA.

Although Planning Analytics for Microsoft Excel supports processing outside of VBA, it is not recommended due to certain environmental constraints. If you need to process outside of VBA, you should start with using the sample files. Processing outside of VBA will require you to make self-service changes and have the system knowledge required to allow for complex usage scenarios.

## Processing within VBA

> The following example demonstrates how to call the Logon method within VBA

```vb
Dim bResult as Boolean

bResult = CognosOfficeAutomationObject.Logon
("http://localhost/ibmcognos/cgi-bin/cognos.cgi","Administrator",
"CognosAdmin", "Production")

If bResult Then

    CognosOfficeAutomationObject.ClearAllData()

    CognosOfficeAutomationObject.RefreshAllData()

    CognosOfficeAutomationObject.Logoff()

    Dim sTraceLog as String 

    sTraceLog = CognosOfficeAutomationObject.TraceLog

    'Here is where you could write the trace log to file.

    MsgBox sTraceLog

End If
```
## Troubleshooting issues

You may encounter issues when processing outside of VBA. This section outlines a few commong issues.

### The script runs, but nothing happens

> Example

```vb
' This example is placed in the CognosOfficeAutomationExample.bas file

Public Sub ClearAllData()
On Error GoTo HANDLER:
Dim oMessageSuppressor As CognosOfficeMessageSuppressor

    'Use the message suppressor to turn off all Cognos Office messages.
    Set oMessageSuppressor = New CognosOfficeMessageSuppressor
    
    Dim test As Object
    Set test = Reporting
    
    'Call the Cognos Office Automation object to clear the data.
    CognosOfficeAutomationObject.ClearAllData
    
    Exit Sub
HANDLER:
    '<Place error handling here. You may not want to display a message box if you are running in a scheduled task>
End Sub
```

If your script runs, but nothing happens, it may be a sign that the Planning Analytics for Microsoft Excel add-in is not activated. If the add-in is not activated, Microsoft Excel will not know when and where to execute the Planning Analytics for Microsoft Excel APIs.


You can activate the Planning Analytics for Microsoft Excel manually by opening the IBM Planning Analytics tab:

1. Launch Microsoft Excel.
2. Click the IBM Planning Analytics tab.

You can also activate the Planning Analytics for Microsoft Excel by adding the following lines before a macro call:

`Dim test As Object`

`Set test = Reporting`


### Application crashes on when script is running

> Example

```vb
' This example is placed in the Automate_COI.vbs file

Dim bSuccess
Dim objExcel
Dim objWB

Set objExcel = CreateObject("Excel.Application")
    
    If objExcel Is Nothing Then WScript.Quit(0)

`Turn off Excel features.
objExcel.Visible = False
objExcel.ScreenUpdating = False
objExcel.DisplayAlerts = False
        
```

If the application crashes when you run your script, you may need to check the Microsoft Excel events before you call the Planning Analytics for Microsoft Excel APIs. 

To check the Microsoft Excel events, add the following lines to your VBS or VBA file:

`objExcel.Visible = False`

`objExcel.ScreenUpdating = False`

`objExcel.DisplayAlerts = False`


You may also need to add a sleep event before calling the Planning Analytics for Microsoft Excel API. 

Example of a sleep event: 

`WScript 5000`