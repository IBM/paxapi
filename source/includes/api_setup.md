# Set up

The quickest way to set up automation is to import the `CognosOfficeAutomationExample.bas` file into the Microsoft application.

This file contains all the necessary macros, including the CognosOfficeAutomationObject macro. Alternatively, you can create templates that contain this .bas file to supply the code for logging on to IBM® Cognos® application, refreshing the content of specified workbooks, documents, or presentations, and logging off.

After the reference to IBM Cognos automation is established, any macro in VBA can call the functions exposed in the IBM Cognos automation API.

If the Microsoft application is open when a command is executing, the command executes in interactive mode. If the Microsoft application is closed when the command is executing, the command executes in batch mode. Executing in batch mode means that all display alerts are turned off.

Because the object is obtained at run time and there is no type library installed on the client machine, you cannot use IntelliSense to determine what properties and methods are available on the object.

## Before you begin

To use the IBM Cognos automation macro files, you must import the `CognosOfficeMessageSuppressor.cls` file. The .cls file contains the SuppressMessages function that allows you to disable the standard alerts and messages.

## Procedure

1. Open a new Office document, workbook, or presentation.
2. Customize the ribbon to display the **Developer** tab.
3. Click the **Developer** tab, and then click **Visual Basic**.
4. Do the following based on the Microsoft Office application you are using:
  * For Microsoft Excel and Microsoft PowerPoint, right-click **VBAProject** and click **Import File**.
  * For Microsoft Word, right-click **Project** and click **Import File**.
The Import File dialog box appears.
5. Browse to the location where the IBM Cognos Automation macro files are installed.
The default location is `<client_installation_directory>\Automation`.
6. For Microsoft Excel or Microsoft Word click the `CognosOfficeAutomationExample.bas` file or for Microsoft PowerPoint click the CognosOfficeAutomationPPExample.bas file and import it into the VBA project.
Do not edit this code module. Do not import both files, which are application specific. This will cause problems for the Open routine.
7. Repeat steps 3 to 5 to import the `CognosOfficeMessageSuppressor.cls` file.
8. Close the Visual Basic Editor and return to the IBM Cognos application.
9. Save the file as a template, close it, and then reopen the template file.

## Results
You can now call the macros contained in the Cognos automation macro files from the VBA code that you write in Excel, Word, or PowerPoint.
