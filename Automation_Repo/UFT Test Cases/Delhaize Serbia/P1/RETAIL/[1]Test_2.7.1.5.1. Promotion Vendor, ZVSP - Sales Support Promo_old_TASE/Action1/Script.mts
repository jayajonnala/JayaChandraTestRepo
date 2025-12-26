
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.7.1.5.1. Promotion Vendor, ZVSP - Sales Support Promo_old
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.7.1.2.2. Create Low Level Header ZSSP - Sales Support Promo
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.7.1.5.1. Promotion Vendor, ZVSP - Sales Support Promo_old"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.2. Create Low Level Header ZSSP - Sales Support Promo.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM) 
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
''----------------------Tcode WAK2----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

'Enter promotion
Call SetTextbox("Promotion","WAKHD-AKTNR","",DT_WAK2_1100_PROMOTION,FALSE)
Call TakeScreenShot()
Call PressEnter() 

'Navigate to Pricing conditions \(purchasing\) Tab
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Pricing conditions (purchasing)",False)
Wait(1)
Call TakeScreenShot()
 
'Click on Conditions
Call ClickButton("Create condition",False) 
Wait(5)
Call TakeScreenShot()

' SetVerticalScrollBar(ScrollBarPosition, blnIsItPopup)
Call SetVerticalScrollBar(39,True)

Call ClickLabel("Article Info Record \(NEW\)",0,True)
'Call ClickLabel("Article Info Record (NEW)",0,True)
'Click on Choose
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Wait(1)

Call SetTextbox("Purch\. Organization","KOMG-EKORG","",DT_WAK2_1906_PURCH_ORGANIZATION,FALSE)
Call SetTextbox("Info record category","KOMG-ESOKZ","",DT_WAK2_1906_INFO_RECORD_CATEGORY,FALSE)

'Enter the details in Table
'Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Vendor",1,DT_WAK2_1906_TABLECELL_VENDOR_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Supplier",1,DT_WAK2_1906_TABLECELL_VENDOR_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_WAK2_1906_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_WAK2_1906_TABLECELL_AMOUNT_0,False)
Call TakeScreenShot()
Call PressEnter() 

Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_FAST_ENTRY",DT_WAK2_1906_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Enter   \(F5\)",True)
wait(2)
Call TakeScreenShot()

Call SetVerticalScrollBar(39,True)
Call ClickLabel("Article Info Record \(NEW\)",1,True)
'Call ClickLabel("Article Info Record (NEW)",1,True)
'Click on Choose
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Wait(1)

Call SetTextbox("Purch\. Organization","KOMG-EKORG","",DT_WAK2_1906_PURCH_ORGANIZATION_OCC2,FALSE)
Call SetTextbox("Info record category","KOMG-ESOKZ","",DT_WAK2_1906_INFO_RECORD_CATEGORY_OCC1,FALSE)

'Enter the details in Table
'Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Vendor",1,DT_WAK2_1906_TABLECELL_VENDOR_0_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Supplier",1,DT_WAK2_1906_TABLECELL_VENDOR_0_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_WAK2_1906_TABLECELL_ARTICLE_0_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_WAK2_1906_TABLECELL_AMOUNT_0_OCC1,False)
Call TakeScreenShot()
Call PressEnter() 
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_FAST_ENTRY",DT_WAK2_1906_CHECK_TEXT_OF_TABLECELL_AMOUNT_0_OCC1)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Enter   \(F5\)",True)
wait(2)
Call TakeScreenShot()

'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
wait(2)
Call TakeScreenShot()

'Verify table Cell Cntent
Call VerifyTableCellContent(1,"Variable Key","SAPMWAKAUEB_EKKO",Lcase(DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_VARIABLE_KEY_0))
Call VerifyTableCellContent(5,"Variable Key","SAPMWAKAUEB_EKKO",Lcase(DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_VARIABLE_KEY_4))
'Call VerifyTableCellContent(3,"Variable Key","SAPMWAKAUEB_EKKO",DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_VARIABLE_KEY_2)
'Call VerifyTableCellContent(7,"Variable Key","SAPMWAKAUEB_EKKO",DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_VARIABLE_KEY_6)
Call VerifyTableCellContent(3,"Amount","SAPMWAKAUEB_EKKO",DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_AMOUNT_2)
Call VerifyTableCellContent(7,"Amount","SAPMWAKAUEB_EKKO",DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_AMOUNT_6)

'Select Check Box
'Call SelectCheckBoxTableByRefColumn("SAPMWAKAUEB_EKKO","S","Variable Key",DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_VARIABLE_KEY_2,"ON")
'Call SelectCheckBoxTableByRefColumn("SAPMWAKAUEB_EKKO","S","Variable Key",DT_WAK2_8218_CHECK_TEXT_OF_TABLECELL_VARIABLE_KEY_6,"ON")

'Save  the Promotion Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBar(DT_WAK2_1100_CHECK_TEXT_OF_STATUSBAR)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


'SAPGuiSession("Session_2").SAPGuiWindow("Valid Condition Types").VerticalScrollbarPosition 19
'SAPGuiSession("Session_2").SAPGuiWindow("Valid Condition Types").SAPGuiLabel("Article Info Record (NEW)").SetFocus
'SAPGuiSession("Session_2").SAPGuiWindow("Valid Condition Types").SAPGuiLabel("Article Info Record (NEW)").SetCaretPos 8
'SAPGuiSession("Session_2").SAPGuiWindow("Valid Condition Types").SendKey ENTER

