'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AP Bill of Exchange_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th March
'.................Modified By :
'.................Modified Date/Details :
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

gstrTestCaseName = "Test_Manage AP Bill of Exchange_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AP Bill of Exchange_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FF67_0101_STATEMENT_NUMBER",(Cint(DT_FF67_0101_STATEMENT_NUMBER)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode FF67----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Statement Date","FEBMKA-AZDAT","",Replace(DT_FF67_0101_STATEMENT_DATE,"/","."),False)
Call SetTextbox("Statement Number","FEBMKA-AZNUM","",DT_FF67_0101_STATEMENT_NUMBER,False)
Call SetTextbox("Posting date","FEBMKA-BUDTM","",Replace(DT_FF67_0101_POSTING_DATE,"/","."),False)
Call SetTextbox("Closing Balance","FEBMKA-ESALD","",DT_FF67_0101_CLOSING_BALANCE,False)
Call SetTextbox("Opening Balance","FEBMKA-SSALD","",DT_FF67_0101_OPENING_BALANCE,False)

'''following fields are obsulute in SAP screen
''Call SetTextbox("Account ID","nnnn","",DT_FF67_0101_ACCOUNT_ID,False)
''Call SetTextbox("House bank","nnnn","",DT_FF67_0101_HOUSE_BANK,False)
''Call SetTextbox("Company Code","nnnn","",DT_FF67_0101_COMPANY_CODE,False)

'''following field is new addition
Call SetTextbox("Bank Account","FEBMKA-BANKN","",DT_FF67_0101_BANK_ACOUNT,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Line item text","FEBMKA-VGMAN","0",DT_FF67_8000_TRAN,False)
Call SetTextbox("Value date","FEBEP-VALUT","0",Replace(DT_FF67_8000_VALUE_DATE,"/","."),False)
Call SetTextbox("Amount","FEBMKA-KWBTR","0",DT_FF67_8000_AMOUNT,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
VerifyStatusBar(DT_FF67_0101_CHECK_TEXT_OF_STATUSBAR)

'this is the replacement of " Call ClickButton("Post statement \(Ctrl\+S\)",False)
Call SelectMenuBar("Bank statement;Post;Individual statement")

wait(1)
Call TakeScreenShot()

Call VerifyWindowValue(DT_FF67_0120_CHECK_TEXT_OF_TITL)

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


