

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.2.1.1.1. Create Vendor Account Group ZVOT - One Time Vendor
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

gstrTestCaseName = "Test_2.2.1.1.1. Create Vendor Account Group ZVOT - One Time Vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''----------------------Tcode XK01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,False)
Call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,False)
Call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,False)
Call TakeScreenShot()
Call PressEnter()  
Wait(2)

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call FocusTextBox("Recon\. account","LFB1-AKONT",False)
'Call SendKey("{F4}")
'Wait(1)
Set objWsh = CreateObject("WScript.Shell") 
objWsh.SendKeys "{F4}" 
Set objWsh=nothing
Wait(4)

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickLabel(DT_XK01_GL_ACCOUNT,"",True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Pmnt Terms","LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)

Call ClickButtonIfExist("Next screen   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call ClickButton("Yes",True)
Wait(3)

Call GetStatusBar("item1","DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
