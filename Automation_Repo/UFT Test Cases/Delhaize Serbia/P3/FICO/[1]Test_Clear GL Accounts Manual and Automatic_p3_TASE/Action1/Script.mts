
''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear GL Accounts Manual and Automatic_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear GL Accounts Manual and Automatic_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\Retail\DT_Intracompany Store returns to DC - SW31_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()   

'''----------------------Tcode F-03----------------------------
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_F03_0131_DOCUMENT_NUMBER,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F03_0131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",Replace(DT_F03_0131_CLEARING_DATE,"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F03_0131_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F03_0131_COMPANY_CODE,False)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F03_0131_CURRENCY,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

'set filter criteria
Call SetTextbox("From","RF05A-SEL01",0,DT_F03_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F03_0731_FROM_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(5)
'Capture the screenshot

'verify table data
Call VerifyTableCellContent(1,"RSD Gross","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL______RSD_GROSS_0)
Call VerifyTableCellContent(2,"RSD Gross","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL______RSD_GROSS_1)
Call VerifyTableCellContent(1,"Document Number","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0)
Call VerifyTableCellContent(2,"Document Number","SAPDF05XTC_6103",DT_F03_6103_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_1)

Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_F03_6103_CHECK_TEXT_OF_NOT_ASSIGNED,False)

'Call ClickButton("Select All",False) 
''Capture the screenshot
'Call TakeScreenShot()
'
'Call ClickButton("Activate Items",False)
''Capture the screenshot
'Call TakeScreenShot()
'
'Call FocusTextBox("Number of items","RF05A-ANZPO",False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
wait(1)
'Capture the screenshot
Call TakeScreenShot()
'Validate If doc is generated
Call GetStatusBar("item1","DT_F03_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F03_0131_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

 


