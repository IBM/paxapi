# Script files

This site includes sample script files that you can use to automate functions. The samples include script files for scheduling the refresh of documents. Also, there is a script file to update the server URL.

You must modify the script files to meet your particular needs or use them as a reference to create your own programs. For more information, see the comments in the file.

These Visual Basic Scripts (VBS) are provided as sample programs and are [here](https://github.com/IBM/paxapi/raw/master/attachments/Automation.zip):

* Automate_COI.vbs
* Automate_COI_Excel.vbs
* UpdateServerURLSample.vbs

## Scripted deployment of single .xll add-in as an Excel persistent add-in

Additionally, [here](https://github.com/IBM/paxapi/raw/master/attachments/scripted_registration_sample.zip) is a script bundle that demonstrates how to pre-register (and unregister) the single-file version of Planning Analytics for Microsoft Excel, it may be helpful in conjunction with centralized software deployment solutions.  The sample includes pre-activating the addin for all user profiles present on the machine, when placed adjacent to the addin xll location and run with elevation.  Current user pre-registration does not require elevation of privileges and is also covered by the sample by modifying the optional parameters.
