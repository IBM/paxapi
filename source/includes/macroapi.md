# Macro files

The macro files for Cognos® Office are written in Microsoft Visual Basic for Applications (VBA).

The files are installed with IBM® Cognos Office in the automation folder. The default location is `[installation_directory]\Automation`.

### Macros
The following macro files are installed.

File | Description
--------- | ------- | -----------
CognosOfficeAutomationExample.bas | Because it is a BASIC file created using VBA, this file has the extension .bas. It contains the `CognosOfficeAutomationObject` property that enables IBM Cognos Office automation in the current document. It also contains wrapper functions that call the API exposed by IBM Cognos Office.
CognosOfficeMessageSuppressor.cls | This file shows how to use the [SuppressMessages](#suppressmessages) API function.