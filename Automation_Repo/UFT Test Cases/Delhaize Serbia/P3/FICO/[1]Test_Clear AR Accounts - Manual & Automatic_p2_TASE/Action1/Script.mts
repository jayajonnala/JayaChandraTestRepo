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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : DT_Clear AR Accounts - Manual & Automatic_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "DT_Clear AR Accounts - Manual & Automatic_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AR Accounts - Manual & Automatic_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode F-32----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F32_0131_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F32_0131_COMPANY_CODE,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",Replace((DT_F32_0131_CLEARING_DATE),"/","."),False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F32_0131_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("From","RF05A-SEL01","0",DT_F32_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F32_0731_FROM_OCC1,False)

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(2)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_F32_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F32_0131_CHECK_TEXT_OF_STATUSBAR)

Call SelectMenuBar("Document;Display")

'verify the grid coponents
call VerifyGridCellContent("",1,"Posting Key","",DT_F32_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_F32_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",1,"Amount","",DT_F32_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

call VerifyGridCellContent("",2,"Posting Key","",DT_F32_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",2,"Account","",DT_F32_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",2,"Amount","",DT_F32_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

