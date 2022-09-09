# Necessary IBM Cognos automation API references
> m_o Cafe object reference

```vb
    Dim m_oCAFE As Object
```

> m_oCOAutomation object reference

```vb
    Private m_oCOAutomation As Object
```
> CognosOfficeAutomationObject() Property Get statement

```vb
    'Returns the instance of the Cognos Office Automation Object.
    Public Property Get CognosOfficeAutomationObject()
    On Error GoTo Handler:

        'Fetch the object if we don't have it yet.
        If m_oCOAutomation Is Nothing Then
            Set m_oCOAutomation = Application.COMAddIns("CognosOffice12.Connect").Object.AutomationServer

            
        End If
    
        Set CognosOfficeAutomationObject = m_oCOAutomation

        Exit Property
    Handler:
        '<Place error handling here.  Remember you may not want to display a message box if you are running in a scheduled task>
    End Property
    Copy
```

If Planning Analytics for Excel and Cognos Microsoft Office are installed together,  Planning Analytics for Excel detects Cognos Microsoft Office and changes the process id from “CognosOffice12.Connect” to “CognosOffice12.ConnectPAfEAddin”.

> CognosOfficeAutomationObject() Property Get statement when Cognos Microsoft Office is installed

```vb 
    'Use the following to adjust your existing VBA scripts when Planning Analytics for Excel and Cognos Microsoft Office are installed together:


    'Returns the instance of the Cognos Office Automation Object.
    Public Property Get CognosOfficeAutomationObject()
    On Error GoTo Handler:

        'Fetch the object if we don't have it yet.
        If m_oCOAutomation Is Nothing Then
            Set m_oCOAutomation = Application.COMAddIns("CognosOffice12.ConnectPAfEAddin").Object.AutomationServer
            
        End If
    
        Set CognosOfficeAutomationObject = m_oCOAutomation

        Exit Property
    Handler:
        '<Place error handling here.  Remember you may not want to display a message box if you are running in a scheduled task>
    End Property
 Copy
```



> Reporting() Property Get statement

```vb
    'Returns the instance of the Cognos Office Automation Object.
    Public Property Get Reporting()
    On Error GoTo Handler:

       'Fetch the object if we don't have it yet.
       If m_oCAFE Is Nothing Then
           Set m_oCAFE = CognosOfficeAutomationObject.Application("COR", "1.1")
       End If
       
       Set Reporting = m_oCAFE

       Exit Property
    Handler:
       MsgBox "Error"
       '<Place error handling here.  Remember you may not want to display a message box if you are running in a scheduled task>
    End Property
```
 


The references mentioned in this section can be imported via the CognosOfficeAutomationExample.bas file. It is good practice to double-check that the file contains all of references. If the CognosOfficeAutomationExample.bas file is missing any references, you can add these references to the file yourself.

<aside class="notice">
Ensure that both object references are initialized before using IBM Planning Analytics for Microsoft Excel API.
</aside>


![alt text](images/api_references.jpg "CognosOfficeAutomationExample.bas file")