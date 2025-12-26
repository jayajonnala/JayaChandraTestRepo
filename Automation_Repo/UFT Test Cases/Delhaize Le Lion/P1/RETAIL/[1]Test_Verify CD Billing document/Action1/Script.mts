
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Verify CD Billing document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'------------------------ME23N---------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
call ClickButton("Other Document.*",False)

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)

'''Call ClickButtonToolBar("&FIND",1)
'''CALL ClickButtonToolBar("shell\[0\]",3)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ME23N_0841_SEARCH_TERM,TRUE)  
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Wait(1)
'
''Call VerifyTextBoxContent("Information Message","GS_SEARCH-SEARCH_INFO","",DT_ME23N_0841_CHECK_TEXT_OF_GS_SEARCHSEARCH_INFO,True)
'call ClickButton("Cancel   \(F12\)",False)
'Call TakeScreenShot
'
''Call GetCellDataGuiGridPopupByRefTwoColumns("","","Cell Content","Group Description","Article Document","Group Description","Article Document","DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VALUE_OUTPUT")
'
'' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)

Call GetGridContent("",0,"Article Document","","Short Text","RE-L","DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VALUE_OUTPUT")
Call ClickButtonIfExist("Close window   \(Enter\)",True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'-------------------------MRI4--------------------------

Call SetTcode(DT_ME23N_0014_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_ME23N_6150_INVOICE_DOCUMENT_NO,FALSE)   
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_ME23N_6150_FISCAL_YEAR,FALSE)   
call ClickButton("Display Document   \(F2\)",False)
Call TakeScreenShot


'Verifying reference text box not required according to TAO log
'CALL VerifyTextBoxContent("Reference","INVFO-XBLNR",3,DT_ME23N_0010_CHECK_TEXT_OF_REFERENCE,FALSE)


Call SetTcode(DT_ME23N_6000_OKCD)

Call LogOff()
Call FinalStatus ()









