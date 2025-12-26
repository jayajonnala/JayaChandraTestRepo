'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'.................Test Script Name : Test_2.3.1.1.1. Create Customer with Account Group ZCCO - DG Competit_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 18th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.3.1.1.1. Create Customer with Account Group ZCCO - DG Competit_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_2.3.1.1.1. Create Customer with Account Group ZCCO - DG Competit_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode XD01----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetComboByKey("RF02D-KTOKD",DT_XD01_7100_ACCOUNT_GROUP)
Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD01_7100_COMPANY_CODE,True)
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD01_7100_SALES_ORGANIZATION,True)
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD01_7100_DISTRIBUTION_CHANNEL,True)
Call SetTextbox("Division","RF02D-SPART","",DT_XD01_7100_DIVISION,True)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XD01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XD01_0301_SEARCH_TERM_12,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT2","",DT_XD01_0301_SEARCH_TERM_12_OCC1,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XD01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XD01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XD01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XD01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XD01_0301_COUNTRY,False)
Call SetTextbox("Region","ADDR1_DATA-REGION","",DT_XD01_0301_REGION,False)
Call SetComboByKey("ADDR1_DATA-LANGU",DT_XD01_0301_LANGUAGE)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Control Data",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Marketing",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Contact Person",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Company Code Data   \(Ctrl\+F2\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Sales Area Data   \(Ctrl\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(3)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_GETMESSAGE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_EXPECTED)

Call ClickButtonIfExist("Cancel   \(F12\)",True)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

