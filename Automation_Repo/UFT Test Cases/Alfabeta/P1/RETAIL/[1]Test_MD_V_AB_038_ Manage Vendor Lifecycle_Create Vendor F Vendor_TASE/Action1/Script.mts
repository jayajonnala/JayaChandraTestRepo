

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_V_AB_038_ Manage Vendor Lifecycle_Create Vendor F Vendor
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

gstrTestCaseName = "Test_MD_V_AB_038_ Create Vendor F Vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  XK01------------------------------------------------
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,False)
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,False)
Call SetTextboxNoLabel("RF02K-LIFNR",0,DT_XK01_0100_VENDOR,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,False)
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 

Call SetTextbox("Recon\. account","LFB1-AKONT","",DT_XK01_0210_RECON_ACCOUNT,False)
Call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,False)
Call SetTextbox("Cash mgmnt group","LFB1-FDGRV","",DT_XK01_0210_CASH_MGMNT_GROUP,False)
Call TakeScreenShot()
Call PressEnter()
Call SetTextboxNoLabel("LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR )

Call LogOff()
Call FinalStatus ()



