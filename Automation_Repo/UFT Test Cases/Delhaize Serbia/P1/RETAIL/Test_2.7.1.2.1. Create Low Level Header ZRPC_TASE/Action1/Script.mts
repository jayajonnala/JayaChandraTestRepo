
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.7.1.2.1. Create Low Level Header ZRPC
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
'.................Test Script Name :Test_2.7.1.2.1. Create Low Level Header ZRPC
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.7.1.2.1. Create Low Level Header ZRPC"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode WAK1----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

'Enter promotion type
Call SetTextbox("Promotion type","WAKHD-AKART","",DT_WAK1_1000_PROMOTION_TYPE,FALSE)
Call TakeScreenShot()

'Click on Fast Entry
Call ClickButton("Fast entry   \(F5\)",False) 
Wait(1)

'Verify the Screen
Call CheckScreen(DT_SAPTRANSACTIONCODE,"Create promotion: Fast Entry Overview")

'Enter the details and click on description Button
Call SetTextbox("Promotion","WAKTD-AKTKT","",DT_WAK1_0500_PROMOTION,FALSE)
Call SetTextbox("Plnd crrncy","WAKHD-WAELA","",DT_CURRENCY,FALSE)
Call TakeScreenShot()

'Click on DEscription
Call ClickButton("Lang\.-dependent description",False) 
Wait(1)

'Verify the Language
Call VerifyTableCellContent(1,"Language","SAPMWAKATEXTE",DT_WAK1_5500_CHECK_TEXT_OF_TABLECELL_LANGUAGE_0)

'Enter the details in Table
Call SetTableDataNoRef("SAPMWAKATEXTE","Language",2,DT_WAK1_5500_TABLECELL_LANGUAGE_1,False)
Call SetTableDataNoRef("SAPMWAKATEXTE","Promotion Descr.",1,DT_WAK1_5500_TABLECELL_PROMOTION_DESCR_0,False)
Call SetTableDataNoRef("SAPMWAKATEXTE","Promotion Descr.",2,DT_WAK1_5500_TABLECELL_PROMOTION_DESCR_1,False)
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

Call SetTextbox("Plnd crrncy","WAKHD-WAELA","",DT_CURRENCY,FALSE)
Call SetTextbox("On sale from","WAKHD-VKDAB","",ConvertDate(DT_WAK1_0500_ON_SALE_FROM),FALSE)
Call SetTextbox("to","WAKHD-VKDBI","",ConvertDate(DT_WAK1_0500_TO),FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Click on Header Data
Call ClickButton("Promotion header data   \(Ctrl\+F1\)",False) 
Wait(1)
Call TakeScreenShot()


Call SetTextbox("Media Type","WAKHD-MEDIA_TYPE","",DT_WAK1_8103_MEDIA_TYPE,FALSE)
Call SetTextbox("Assigned to Promo\.","WAKHD-REF_PROMO","",DT_WAK1_1000_ASSIGNED_TO_PROMO,FALSE)
Call SetTextbox("Purchasing Org\.","WAKHD-EKORG","",DT_WAK1_8230_PURCHASING_ORG,FALSE)
Call TakeScreenShot()

'Navigate to Control Tab
Call SelectTab("TAXI_TABSTRIP_KOPFDATEN","Control",False)
Wait(1)
Call TakeScreenShot()

'Enter the details
Call SetTextbox("Cond\. group purch\.","WAKHD-KOBOG","",DT_WAK1_8231_COND_GROUP_PURCH,FALSE)
Call SetCombo("WAKHD-KALKE","Site group")
Call SetTextbox("Sales Organization","WAKHD-VKORG","",DT_WAK1_8231_SALES_ORGANIZATION,FALSE)
Call SetTextbox("Distribution Channel","WAKHD-VTWEG","",DT_WAK1_8231_DISTRIBUTION_CHANNEL,FALSE)
Call SetTextbox("PP determ\. sequence","WAKHD-EKERV","",DT_WAK1_8231_PP_DETERM_SEQUENCE,FALSE)
Call SelectCheckbox("WAKHD-VZKON",0,DT_WAK1_8231_CREATE_DC_PRICE_CONDITION,False)
Wait(2)
Call SetTextbox("Listing Check","WAKHD-VTART","","",FALSE)
Call SelectCheckbox("WAKHD-EKKON",0,DT_WAK1_8231_CREATE_PURCHASE_PRICE_COND,False)
Call SelectCheckbox("WAKHD-VKKON",0,DT_WAK1_8231_CREATE_SALES_PRICE_COND,False)
'Call SelectCheckbox("WAKHD-NO_KALK",0,DT_WTH_CALC,False)
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Navigate to Site Groups Tab
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Site groups",False)
Wait(1)
Call TakeScreenShot()

''Enter the details in Table
Call SetTableDataNoRef("SAPMWAKAUEB_ZWRK","Site group",1,DT_WAK1_8215_TABLECELL_SITE_GROUP_0,False)
Call PressEnter()
Call TakeScreenShot()

'Save  the Promotion Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_PROMOTION_NO_OUTPUT")
VerifyStatusBar("Promotion " & DT_PROMOTION_NO_OUTPUT&" "&"created")

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

