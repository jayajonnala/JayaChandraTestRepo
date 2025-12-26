

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.01.02 Clear AP Accounts (Manual and Automatic)_P3
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.04.03.01.02 Clear AP Accounts (Manual and Automatic)_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.01 Manage Manual Post  Direct Domestic Vendor Invoic.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-F44----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Currency","BKPF-WAERS","",DT_F44_0131_CURRENCY,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F44_0131_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F44_0131_ACCOUNT,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F44_0131_CLEARING_DATE),False)
Call TakeScreenShot

Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("From","RF05A-SEL01",0,DT_F44_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F44_0731_FROM_OCC1,False)
'
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Amount entered", "RF05A-BETRG", "", DT_F44_6102_CHECK_TEXT_OF_AMOUNT_ENTERED, False)
Call VerifyTableCellContent(1,"Document Number","SAPDF05XTC_6102",DT_F44_6102_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0)
Call VerifyTableCellContent(2,"Document Number","SAPDF05XTC_6102",DT_F44_6102_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_1)

Call VerifyTextBoxContent("Amount entered", "RF05A-BETRG", 0, DT_F44_6102_CHECK_TEXT_OF_AMOUNT_ENTERED, False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_3_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_3_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_3_OUTPUT",DT_DOC_3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call LogOff
Call FinalStatus()



