		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.01.03 Manage Manual Customer Invoicing-Crediting
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

gstrTestCaseName = "Test_09.06.01.01.03 Manage Manual Customer Invoicing-Crediting"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-FB08----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot

Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05A-GJAHS", "", Year(DT_FB08_0105_FISCAL_YEAR), False)
Call SetTextbox("Reversal Reason", "UF05A-STGRD", "",DT_FB08_0105_REVERSAL_REASON, False)
Call SetTextbox("Posting Date", "BSIS-BUDAT", "", ConvertDAte(DT_FB08_0105_POSTING_DATE), False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Display document before reversal   \(F5\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''--------TransactionCode-F.62 ----------''''

Call SetTcode(DT_FB08_0105_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("STORBL", 0, DT_FB08_1000_REVERSE_DOCUMENTS, False)
Call SetTextbox("Fiscal year","RGJAHR-LOW","",Year(DT_FB08_1000_FISCAL_YEAR),False)
Call SetTextbox("Company code","RBUKRS-LOW","",DT_FB08_1000_COMPANY_CODE,False)
Call SetTextbox("Document number","RBELNR-LOW","",DT_FB08_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Document type","RBLART-LOW","",DT_FB08_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Correspondence","REVENT","",DT_FB08_1000_CORRESPONDENCE,False)
Call SetTextbox("Delete if finished since","RDELDAYS","","",False)
'Call SetTextbox("Delete if finished since","RDELDAYS","",DT_FB08_1000_DELETE_IF_FINISHED_SINCE,False)
Call SetTextbox("Log to printer","PRDEST","",DT_FB08_1000_LOG_TO_PRINTER,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickBUtton("Execute   \(F8\)",False)
'Call VerifyTextBoxContent("Information Message","MESSTXT1","",DT_FB08_0010_CHECK_TEXT_OF_MESSTXT1,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode-SP02 ----------''''

Call SetTcode(DT_FB08_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

DT_DOC_TITLE_IP = DT_FB08_1000_DOCUMENT_NUMBER&"/"&Year(DT_FB08_1000_FISCAL_YEAR)
Call VerifyifGuiLabelExistsByRelativeid(DT_DOC_TITLE_IP,"wnd\[0\]/usr/lbl\[49,3\]")

Call SetFocusGuiLabel(DT_DOC_TITLE_IP,"347","56", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot
Call VerifyifGuiLabelExistsByRelativeid(DT_DOC_TITLE_IP,"wnd\[0\]/usr/lbl\[49,3\]")
Call TakeScreenshot
Call ClickLabel("X__PDF", "0", False)
'Wait time added for the loading of PDF screen
Wait(10)
Call TakeScreenshot

Call LogOff'
Call FinalStatus()

