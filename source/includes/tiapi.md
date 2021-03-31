# TurboIntegrator functions

## Before you begin

You must use Microsoft Excel 2007 or a later version to have the option to create ActiveX command button controls.

## Procedure

1. In Microsoft Excel, customize the ribbon to show the **Developer** tab.
2. Add an ActiveX command button control to the worksheet.
For more information about creating a command button, see the Microsoft web site.
3. Right-click the command button and click **View Code**.
4. Add `ExecuteFunction` to the command button.

## Results
To use the command button, you must be logged into the TM1 system specified in the ExecuteFunction call. You can use an automation function to log into the TM1 system. To learn more about ExecuteFunction, see ExecuteFunction.

<aside class="notice">
The `installation_location\Automation\COAutomationExample.xls` sample file contains code for `ExecuteFunction`. The `ExecuteFunction` code demonstrates how to use the IBM® Cognos® automation API to execute TurboIntegrator scripts. For information about TurboIntegrator functions, see the TM1 TurboIntegrator guide.
</aside>

## ExecuteFunction

> The following is an example of an ExecuteFunction method, which creates a subset called "TITest" in the "plan_version" dimension:

```vb
Public Sub ExecuteFunction "http://host_address:host_port", 
"Planning Sample", "CreateSubset", "plan_version", "TITest"
On Error GoTo Handler:
Dim oMessageSuppressor As CognosOfficeMessageSuppressor

    'Use the message suppressor to turn off all Cognos Office messages.
    Set oMessageSuppressor = New CognosOfficeMessageSuppressor
    
    Dim s As String
    
    If Not IsMissing(arg1) Then s = arg1
    If Not IsMissing(arg2) Then s = s + "," + arg2
    If Not IsMissing(arg3) Then s = s + "," + arg3
    If Not IsMissing(arg4) Then s = s + "," + arg4

    'Call the Cognos Office Automation object to refresh the data.
    CognosOfficeAutomationObject.ExecuteFunction host, server, 
    processName, s

    Exit Sub
End Sub
```

> Note
You can specify multiple TI process parameters by separating them with commas.



ExecuteFunction is a method used to execute a specified TurboIntegrator (TI) process from your report.

### Syntax

The following string is the syntax for the Create method.

`CognosOfficeAutomationObject.ExecuteFunction(<data source URL>, <server name>, <TI process name>, <Optional TI process parameter>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
datasource URL | URL of the host system which the Custom Report is to be created from. | String
server name | Name of the server which the Custom Report is to be created from. | String
TI process name | Name of the cube which the Custom Report is to be created from. | String
Optional TI process parameter | Name of the view which the Custom Report is to be created from. | String array

<aside class="notice">
You can include an array of TI process parameters by separating them with commas.
</aside>
