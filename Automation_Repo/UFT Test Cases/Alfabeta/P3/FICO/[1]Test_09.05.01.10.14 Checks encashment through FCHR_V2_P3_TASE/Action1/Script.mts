
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.05.01.10.14 Checks encashment through FCHR_V2_P3
'''''''.................Author : TCS 
'''''''................ Creation Date :
'''''''.................Modified By :
'''''''.................Modified Date/Details :
'''''''
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''''''

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.05.01.10.14 Checks encashment through FCHR_V2_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


''''''''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-FCH6 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Paying company code","PAYR-ZBUKR","",DT_FCH6_0600_PAYING_COMPANY_CODE,False)
Call SetTextbox("House bank","PAYR-HBKID","",DT_FCH6_0600_HOUSE_BANK,False)
Call SetTextbox("Account ID","PAYR-HKTID","",DT_FCH6_0600_ACCOUNT_ID,False)
Call SetTextbox("Check number","PAYR-CHECT","",DT_FCH6_0600_CHECK_NUMBER,False)
Call TakeScreenShot
Call ClickButton("Display   \(Shift\+F5\)",False)

Call SetTextbox("Paying company code","PAYR-ZBUKR","",DT_FCH6_0100_PAYING_COMPANY_CODE,False)
Call SetTextbox("House bank","PAYR-HBKID","",DT_FCH6_0100_HOUSE_BANK,False)
Call SetTextbox("Account ID","PAYR-HKTID","",DT_FCH6_0100_ACCOUNT_ID,False)
Call SetTextbox("Check number","PAYR-CHECT","",DT_FCH6_0100_CHECK_NUMBER,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Change/cash   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SetTextbox("Check encashment","PAYR-BANCD","",ConvertDate(DT_FCH6_0601_CHECK_ENCASHMENT),False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(Lcase(DT_FCH6_0101_CHECK_TEXT_OF_STATUSBAR))

Call LogOff'
Call FinalStatus()'
