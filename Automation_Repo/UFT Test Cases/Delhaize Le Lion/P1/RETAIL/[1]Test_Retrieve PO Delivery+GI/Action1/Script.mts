
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


gstrTestCaseName = "Test_Retrieve PO Delivery+GI"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'------------------------------------me23n-------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


call ClickButton("Other Purchase Order.*",False)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
call ClickButton("Other Document.*",False)


Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ME23N_0841_SEARCH_TERM,TRUE) 
Call TakeScreenshot()

Call TakeScreenshot()
'Call VerifyTextBoxContent("Information Message","GS_SEARCH-SEARCH_INFO","",DT_ME23N_0841_CHECK_TEXT_OF_GS_SEARCHSEARCH_INFO,True)
call ClickButton("Cancel   \(F12\)",False)
'Call ClickButtonIfExist("Cancel   \(F12\)",False)
Call TakeScreenshot()

'''GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
Call GetGridContent("",0,"Article Document","","Short Text",DT_ME23N_0841_SEARCH_TERM,"DT_Articledoc_1")


'Call GetGridContentByRefColumn("",0,"Short Text",DT_ME23N_0841_SEARCH_TERM,"Article Document","DT_Articledoc_1")

'CALL ClickButtonToolBar("shell\[0\]",3)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ME23N_0841_SEARCH_TERM_OCC1,TRUE)
Call TakeScreenshot()
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenshot()
'Call VerifyTextBoxContent("Information Message","GS_SEARCH-SEARCH_INFO","",DT_ME23N_0841_CHECK_TEXT_OF_GS_SEARCHSEARCH_INFO,True)
call ClickButton("Cancel   \(F12\)",False)
'Call ClickButtonIfExist("Cancel   \(F12\)",False)
Call TakeScreenshot()
Call GetGridContent("",0,"Article Document","","Short Text",DT_ME23N_0841_SEARCH_TERM_OCC1,"DT_Articledoc_2")
''Call GetGridContent("",0,"Article Document","5","","","DT_Article")
''call GetGridContentByRefColumn("",0,"Short Text",DT_ME23N_0841_SEARCH_TERM_OCC1,"Article Document","DT_Articledoc_2")
'
''Call SetTcode(DT_ME23N_0014_OKCD)
'
''Call PressEnter() 

Call LogOff()
Call FinalStatus ()






