

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article
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


gstrTestCaseName = "Test_01PRI00_002_Create_for_an_Existing_Article"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  ME12------------------------------------------------
'Call SetTextbox("Vendor","EINA-LIFNR","",DT_ME12_100_VENDOR,False)
Call SetTextboxNoLabel("EINA-LIFNR","",DT_ME12_100_VENDOR,False) '-- as vendor is changed to supplier

Call SetTextbox("Article","EINA-MATNR","",DT_ME12_100_ARTICLE,False)
Call SetTextbox("Purchasing Org\.","EINE-EKORG","",DT_ME12_100_PURCHASING_ORG,False)
Call SetTextbox("Site","EINE-WERKS","",DT_ME12_100_SITE,False)
Call TakeScreenShot()

Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Conditions   \(F8\)",False) 
Call TakeScreenShot()

Call ClickButton("New validity period with reference   \(F8\)",True)
Call GetTableCellData("SAPMV13ATCTRL_D0201","Amount",1,"","","DT_ORIGINAL_AMOUNT_OUTPUT",False)
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_D0201","Amount",1,"","",DT_ME12_201_TABLECELL_AMOUNT_0,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBarMessageType("S")

Call LogOff()
Call FinalStatus ()

