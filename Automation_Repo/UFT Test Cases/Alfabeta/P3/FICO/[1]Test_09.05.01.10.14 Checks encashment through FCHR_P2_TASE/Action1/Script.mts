
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.05.01.10.14 Checks encashment through FCHR_P2
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

gstrTestCaseName = "Test_09.05.01.10.14 Checks encashment through FCHR_P2"
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

''''--------TransactionCode-FCH8 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Paying company code","PAYR-ZBUKR","",DT_FCH8_0800_PAYING_COMPANY_CODE,False)
Call SetTextbox("House bank","PAYR-HBKID","",DT_FCH8_0800_HOUSE_BANK,False)
Call SetTextbox("Account ID","PAYR-HKTID","",DT_FCH8_0800_ACCOUNT_ID,False)
Call SetTextbox("Check number","PAYR-CHECT","",DT_FCH8_0800_CHECK_NUMBER,False)
Call SetTextbox("Void reason code","PAYR-VOIDR","",DT_FCH8_0800_VOID_REASON_CODE,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_FCH8_0800_REVERSAL_REASON,False)
Call SetTextbox("Posting Date","RF05R-BUDAT","",ConvertDate(DT_FCH8_0800_POSTING_DATE),False)
Call SetTextbox("Posting period","RF05R-MONAT","",ConvertDoubledigit(CSTR(Month(DT_FCH8_0800_POSTING_PERIOD))),False)

Call TakeScreenShot
Call PressEnter()     
Call VerifyStatusBarMessageType("S")

Call TakeScreenShot
Call ClickButton("Check information   \(Shift\+F8\)",False)
Call SendKey ("+{F8}")
Call ClickButton("Back   \(F3\)",False)
Wait(3)
Call ClickButton("Cancel payment   \(Shift\+F5\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item2","DT_FCH8_0800_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FCH8_0800_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FCH8_0800_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(Lcase(DT_FCH8_0800_CHECK_TEXT_OF_STATUSBAR))

Call LogOff'
Call FinalStatus()
