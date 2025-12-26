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

''Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name      : Test_Manage AR Bill of Exchange_p5_TASE
'.................Author : TCS          :Bitan
'................ Creation Date         : 20th April
'.................Modified By           :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage AR Bill of Exchange_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AR Bill of Exchange_p5_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode F-32----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Currency","BKPF-WAERS","",DT_F32_0131_CURRENCY,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F32_0131_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F32_0131_PERIOD,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDateFormat(DT_F32_0131_CLEARING_DATE),False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F32_0131_ACCOUNT,False)


Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
Call TakeScreenShot()

Call ClickButton("Select All",False) 
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False) 
Call TakeScreenShot()

Call ClickButton("Document Overview   \(Shift\+F2\)",False)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Account","RF05A-NEWKO","",DT_F32_0700_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F32_0700_SGL_IND,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F32_0700_PSTKY,False)
Call TakeScreenShot()

Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Due on","BSEG-ZFBDT","",ConvertDateFormat(DT_F32_2320_DUE_ON),False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F32_2320_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F32_2320_TEXT,False)
Call TakeScreenShot()

Call FocusTextBox("Bank Key","BSEC-BANKL",False)
Call TakeScreenShot()
Call SendKey("{F2}")
wait(5)
Call PressEnter() 
wait(5)
Call SendKey("{ENTER}")
wait(3)
Call FocusTextBox("Drawee","BSED-WBZOG",True)
wait(3)
Call SendKey("{DOWN}")
wait(3)
Call TakeScreenShot()
Call SendKey("{F2}")

Call SetTextbox("Requested line item","\*BSEG-BUZEI","",DT_F32_0610_REQUESTED_LINE_ITEM,True)
Call TakeScreenShot()
'Call ClickButton("Continue Processing   \(Shift\+F1\)",True)
Call SendKey("{ENTER}")
wait(2)

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F32_2320_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F32_2320_ACCOUNT,False)
Call TakeScreenShot()

Call PressEnter() 
wait(3)
Call TakeScreenShot()

Call PressEnter() 
wait(3)
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F32_0301_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F32_0301_BUS_AREA,False)
Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_F32_0301_PAYT_TERMS,False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",ConvertDateFormat(DT_F32_0301_BLINE_DATE),False)
Call TakeScreenShot()

Call SelectMenuBar("Document;Simulate")
Wait(2)
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True)
wait(2)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F32_0131_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()

