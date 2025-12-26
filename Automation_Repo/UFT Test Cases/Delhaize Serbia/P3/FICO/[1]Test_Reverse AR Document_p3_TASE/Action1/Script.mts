

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Reverse AR Document_p3
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

gstrTestCaseName = "Test_Reverse AR Document_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
'----------------------Tcode F-32----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


'Enter the Details
Call SetTextbox("Account","RF05A-AGKON","",DT_F32_0131_ACCOUNT,FALSE)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F32_0131_CLEARING_DATE),FALSE)
Call SetTextbox("Period","BKPF-MONAT","",DT_F32_0131_PERIOD,FALSE)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F32_0131_COMPANY_CODE,FALSE)
Call SetTextbox("Currency","BKPF-WAERS","",DT_F32_0131_CURRENCY,FALSE)

Call SelectRadioButton("RF05A-XPOS1","Document Number",False)
Call TakeScreenShot()
Call PressEnter()


Call SetTextbox("From","RF05A-SEL01",0,DT_F32_0731_FROM,FALSE)
Call SetTextbox("From","RF05A-SEL01",1,DT_F32_0731_FROM_OCC1,FALSE)
Call TakeScreenShot()

'Click on Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(2)
Call PressEnter()
Call PressEnter()

'Verify the Content
''''Call VerifyTableCellContent(1,"RSD Gross","SAPDF05XTC_6102",DT_F32_6102_CHECK_TEXT_OF_TABLECELL_RSD_GROSS_0)
''''Call VerifyTableCellContent(2,"RSD Gross","SAPDF05XTC_6102",DT_F32_6102_CHECK_TEXT_OF_TABLECELL_RSD_GROSS_1)
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB",0,DT_F32_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)


'Click Post
Call ClickButton("Post   \(Ctrl\+S\)",False) 

'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code RS01")

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


