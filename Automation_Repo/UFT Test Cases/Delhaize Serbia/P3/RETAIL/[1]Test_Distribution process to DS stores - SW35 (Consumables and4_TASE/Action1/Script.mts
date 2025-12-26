
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW35 (Consumables and4
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW35 (Consumables and4
'.................Author : TCS 	   
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW35 (Consumables and4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Distribution process to DS stores - SW35 (Consumables and4_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ZMDIM_ZEROING_STOCK----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call SetCombo("P_OPCO",DT_ZMDIM_ZEROING_STOCK_1000_OPERATING_COMPANY)
Call SetTextbox("Material Type","S_MTART-LOW",0,DT_ZMDIM_ZEROING_STOCK_1000_ARTICLE_TYPE,False)
Call SetTextbox("Company Code","S_BUKRS-LOW",0,DT_ZMDIM_ZEROING_STOCK_1000_COMPANY_CODE,False)
Call SetTextbox("Site","S_WERKS-LOW",0,DT_ZMDIM_ZEROING_STOCK_1000_SITE,False)
Call SetTextbox("Storage Location","S_LGORT-LOW",0,DT_ZMDIM_ZEROING_STOCK_1000_STORAGE_LOCATION,False)

Call SetTextbox("Entry Date","S_CPUDT-LOW",0,DT_ZMDIM_ZEROING_STOCK_1000_ENTRY_DATE,False)
Call SetTextbox("Year current period","P_LFGJA",0,DT_ZMDIM_ZEROING_STOCK_1000_YEAR_CURRENT_PERIOD,False)
Call SetTextbox("Period","P_LFMON",0,DT_ZMDIM_ZEROING_STOCK_1000_PERIOD,False)

Call SetCombo("P_FUNC",DT_ZMDIM_ZEROING_STOCK_1000_POSTING_FUNCTIONALITY)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Wait(5)
Call TakeScreenShot()
'

'' '''ClickButton(tooltipOrButtonName, blnIsItPopup)
'Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDIM_ZEROING_STOCK_SEARCH_TERM,True)
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD","0",DT_ZMDIM_ZEROING_STOCK_FIND_ONLY_ENTIRE_WORD_OR_VALUE,True)
'Call SetComboByKey("GS_SEARCH-SEARCH_ORDER",DT_ZMDIM_ZEROING_STOCK_SEARCH_DIRCT)
'Call TakeScreenShot
'Call PressEnter()
'Call ClickButton("OK   \(Enter\)",True)
'Call ClickButton("Cancel   \(F12\)",True)
'Call ClickButton("Add Filter Criterion \(F7\)",True)
'Call ClickButton("Define Filter Values",True)
'Call TakeScreenShot
'Call SetTextbox("Document number","%%DYN001-LOW","",DT_ZMDIM_ZEROING_STOCK_1105_DOCUMENTNUMBER,True)
'Call TakeScreenShot
'Call ClickButton("Execute   \(Enter\)",True)
'Call TakeScreenShot
'
Call SelectRowGuiGrid("",0,"Article",DT_ARTICLE_REFNO,False)
Wait(1)
Call TakeScreenShot()

Call ClickButton("Select all   \(F5\)",False) 
Wait(1)
Call TakeScreenShot()

Call ClickButton("Save Documents   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Click on Details Button
Call ClickButton("Details   \(Ctrl\+Shift\+F3\)",False) 
Wait(1)
Call TakeScreenShot()

Call VerifyGridCellContent("",22,"Cell Content...",0,DT_ZMDIM_ZEROING_STOCK_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_0_T_MSG)

 
Call ClickButtonIfExist("Close window   \(Enter\)",True)
wait(2)

'Call GetGridContent("",0,"Document number",1,"Article","502348",DT_ZEROING_STOCK_DOC_OUTPUT)
'Call GetGridContentByRefColumn("",0,"Article","502348","Document number",DT_ZEROING_STOCK_DOC_OUTPUT)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

