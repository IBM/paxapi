# REST API
You can use REST APIs to communicate with the TM1 Server.

Before you begin, make sure that you've returned the connection object. The connection object will allow you to implement the REST request methods (GET, POST, DELETE, PATCH) and communicate with the TM1 Server.

To learn more about the connection object, see [GetConnection](#getconnection).

## GET requests

> Example of how you can use a `GET` request in a VBA module to return a JSON object.

```vb
Public Property Get OData(Server As String) As String
    'OData endpoint
    OData = "tm1/" + Server + "/api/v1"
End Property

Public Property Get Current() As String
    Current = Reporting.Settings.GetValue("MruServer")
End Property

Public Function GetCurrentServer() As String
   Dim sServerCubeMRU As String
   Dim sServerCube As Variant
   Dim sServer As String
   sServerCubeMRU = Reporting.Settings.GetValue("MruPackage")
   sServerCubeMRU = Mid(sServerCubeMRU, 2, Len(sServerCubeMRU) - 2)
   sServerCube = Split(sServerCubeMRU, ",")
   sServer = Split(sServerCube(0), ":")(1)
   GetCurrentServer = Mid(sServer, 3, Len(sServer) - 3)
End Function

Public Function oDataGet(path As String) As JSONObjectWrapper
    Dim result As New JSONParser
    Dim response As Object
    Set response = Reporting.GetConnection(Current).Get(OData(GetCurrentServer) & "/" & path)
End Function
```
Use `GET` requests to return data from the TM1 Server.

### Syntax

The following string is the syntax for the `GET` request.

`Reporting.GetConnection(<CURRENT>).Get(<PATH>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
CURRENT | The URL of the host that you are using the GET request on. | String
PATH | The full (absolute) path that you are using the GET request on. | String

## POST requests

> Example of how you can use a `POST` request in a VBA module to update a component.

```vb
Public Property Get HubEndpoint() As String
    'Endpoint that CAFE connects to
    HubEndpoint = "pmhub/pm/tm1/"
End Property

Public Property Get OData(Server As String) As String
    'OData endpoint
    OData = "tm1/" + Server + "/api/v1"
End Property

Public Property Get Current() As String
    Current = Reporting.Settings.GetValue("MruServer")
End Property

Public Function GetCurrentServer() As String
   Dim sServerCubeMRU As String
   Dim sServerCube As Variant
   Dim sServer As String
   sServerCubeMRU = Reporting.Settings.GetValue("MruPackage")
   sServerCubeMRU = Mid(sServerCubeMRU, 2, Len(sServerCubeMRU) - 2)
   sServerCube = Split(sServerCubeMRU, ",")
   sServer = Split(sServerCube(0), ":")(1)

   GetCurrentServer = Mid(sServer, 3, Len(sServer) - 3)
End Function

Public Function oDataPost(path As String, payload As String) As JSONObjectWrapper
    Dim result As New JSONParser
    Dim response As Object
    Set response = Reporting.GetConnection(Current).Post(OData(GetCurrentServer) & "/" & path, payload)
End Function
```
Use `POST` requests to store or update components in the TM1 Server.

### Syntax

The following string is the syntax for the `POST` request.

`Reporting.GetConnection(<CURRENT>).Post(<PATH>, <PAYLOAD>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
CURRENT | The URL of the host that you want to store or update on. | String
PATH | The full (absolute) path of the component that you want to store to or update. | String
PAYLOAD | The JSON payload that you are storing or updating to the TM1 Server. | String

## DELETE requests

> Example of how you can use a `DELETE` request in a VBA module to delete data.

```vb
Public Property Get HubEndpoint() As String
    'Endpoint that CAFE connects to
    HubEndpoint = "pmhub/pm/tm1/"
End Property

Public Property Get OData(Server As String) As String
    'OData endpoint
    OData = "tm1/" + Server + "/api/v1"
End Property

Public Property Get Current() As String
    Current = Reporting.Settings.GetValue("MruServer")
End Property

Public Function GetCurrentServer() As String
   Dim sServerCubeMRU As String
   Dim sServerCube As Variant
   Dim sServer As String
   sServerCubeMRU = Reporting.Settings.GetValue("MruPackage")
   sServerCubeMRU = Mid(sServerCubeMRU, 2, Len(sServerCubeMRU) - 2)
   sServerCube = Split(sServerCubeMRU, ",")
   sServer = Split(sServerCube(0), ":")(1)

   GetCurrentServer = Mid(sServer, 3, Len(sServer) - 3)
End Function

Public Function oDataDelete(path As String) As JSONObjectWrapper
    Dim result As New JSONParser
    Dim response As Object
    Reporting.GetConnection(Current).Delete(OData(GetCurrentServer) & "/" & path)
End Function
```
Use `DELETE` requests to delete components or data in the TM1 Server.

### Syntax

The following string is the syntax for the `DELETE` request.

`Reporting.GetConnection(<CURRENT>).Delete(<PATH>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
CURRENT | The URL of the host that you want to delete. | String
PATH | The full (absolute) path of the component that you want to delete. | String

## PATCH requests

Use `PATCH` requests to update components in the TM1 Server at a target location.

### Syntax

The following string is the syntax for the `PATCH` request.

`Reporting.GetConnection(<CURRENT>).PATCH(<PATH>, <PAYLOAD>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
CURRENT | The URL of the host that you want to update on. | String
PATH | The full (absolute) path of the component that you want to update. | String
PAYLOAD | The JSON payload that you are storing in the TM1 Server. | String

## PUT requests

Use `PUT` requests to place components in the TM1 Server at a target location.

### Syntax

The following string is the syntax for the `PUT` request.

`Reporting.GetConnection(<CURRENT>).PUT(<PATH>, <PAYLOAD>)`

### Arguments
Argument | Description | Data type
--------- | ------- | -----------
CURRENT | The URL of the host that you want to update on. | String
PATH | The full (absolute) path of the component that you want to place. | String
PAYLOAD | The JSON payload that you are storing in the TM1 Server. | String