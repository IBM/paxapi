---
title: PAx API

language_tabs:
  - vb

toc_footers:
  - <a href='https://www.ibm.com/support/knowledgecenter/SSD29G_2.0.0/kc_gen/com.ibm.swg.ba.cognos.ipa.doc_using_planning_analytics_toc-gen2.html'>Full Documentation Here</a>

includes:
  - api_setup
  - api_automation_error
  - api_automation_references
  - globalapi
  - explorationapi
  - quickreportapi
  - dynamicreportapi
  - customreportapi
  - tiapi
  - restapi
  - scriptapi
  - macroapi
  - exampleapi

search: true
---
# Introduction

 Using an application programming interface (API), you can automate the refreshing or publishing of content.

You can use the API to create a scheduled batch program to refresh content on a daily, weekly, or monthly basis so that, as your period data changes, the affected files are kept up-to-date.

You can call the API within Microsoft Excel workbooks using VBA or using VBS and a command line interface. For these types of automation to work, you must register one or more macros within the workbook.

If you have IBM® Cognos® Office installed, you can also use the API in Microsoft Word and Microsoft PowerPoint.

When using sample macros and script files as part of your own processing functions, remember that the API is accessible only as user defined functions (UDFs). UDFs are functions created in Visual Basic for Applications (VBA). In this case, however, the UDFs are created within the IBM Cognos solution and are called from VBA.

To help you understand what is possible using this API, several samples are provided. You can use the samples to help you create your own solutions.

* Creating VBA macros
* Passing parameters, leveraging VBS and the command line interface

In addition to these capabilities, you can schedule scripts, either ones that you create or the samples, to run as a batch process at a set time.

To use automation, you must set your macro security to an appropriate level in your Microsoft application. You can set the macro security level using one of the following options depending on your version of Microsoft Office.

* Change the security level of your Microsoft application to medium or low.
* Change the trusted publishers setting of your Microsoft application so that installed add-ins or templates are trusted.

## Report issues

Any issues or errors related to the Planning Analytics for Microsoft Excel API documentation can be reported in GitHub (https://github.com/IBM/paxapi/issues).

To report any issues or errors related to Planning Analytics for Microsoft Excel API features or functionality, use the IBM Planning Analytics Community (https://community.ibm.com/community/user/businessanalytics/communities/community-home?CommunityKey=8fde0600-e22b-4178-acf5-bf4eda43146b).

