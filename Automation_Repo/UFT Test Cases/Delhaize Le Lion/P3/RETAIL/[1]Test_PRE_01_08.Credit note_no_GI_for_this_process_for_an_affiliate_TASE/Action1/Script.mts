'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_01_08.Credit note_no_GI_for_this_process_for_an_affiliate 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_01_08.Credit note_noaffiliate"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-VA01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 


Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call TakeScreenShot
Call SetTextbox("Billing Date","VBKD-FKDAT","",ConvertDate(DT_VA01_4415_BILLING_DATE),False)
Call SetComboByKey("VBAK-FAKSK"," ")
Call SetTableData("SAPMV45ATCTRL_U_ERF_GUTLAST", "Article", 1, "", "", DT_VA01_4902_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_GUTLAST", "Target quantity", 1, "", "", DT_VA01_4902_TABLECELL_TARGET_QUANTITY_0, False)
Call TakeScreenShot
Call PressEnter() 
Call PressEnter()
Call SetTableData("SAPMV45ATCTRL_U_ERF_GUTLAST", "Site", 1, "", "", DT_SITE, False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_CREDIT_MEMO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_VA01_4001_CHECK_TEXT_OF_STATUSBAR)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)

'''''''''--------TransactionCode-/nVF04----------''''

Call SetTcode(DT_VA01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SelectCheckbox("P_ALLEL", 0, DT_VA01_8001_DELIVERYRELATED, False)
Call SelectCheckbox("P_ALLEA", 0, DT_VA01_8001_ORDERRELATED, False)
Call SetTextbox("Sales Organization","P_VKORG","",DT_VA01_8001_SALES_ORGANIZATION,False)
Call TakeScreenShot
Call ClickButtonIfExist("Display billing list   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Set filter   \(Ctrl\+Shift\+F2\)",False)
Call SelectRowGuiGridbyRowNo("Column Set", 0, 8, True)
'Call ClickButtonToolBar("&FIND",1)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Sales Document",True)
'Call SetCombo("Search Direction","From Start of Table Downwards")
Call TakeScreenShot
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",1,"ON",True)
'Call ClickButtonIfExist("OK   \(Enter\)",True)
'Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
Call ClickButtonIfExist("Define Filter Values",True)
Call TakeScreenShot
Call SetTextboxNoLabel("%%DYN001-LOW","",DT_VA01_1105_SALES_DOCUMENT,True)
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
'Changes are made due to parameter update.

''Call GetLabelContentByRefLabel("Group", 0,-32,"DT_OP_VA01_0120_CHECK_TEXT_OF_NO_NAME", False)
'Call GetLabelContentByRefLabel("Group",0,-32,"DT_OP_VA01_0120_CHECK_TEXT_OF_NO_NAME", False)
Call GetLabelContentByRefLabel("Group",0,-36,"DT_OP_VA01_0120_CHECK_TEXT_OF_NO_NAME", False)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

