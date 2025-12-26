		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.07.01 NOVA account payment by Cash_Movement Type F504

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


gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.06.05.07.01 NOVA account payment by Cash_Movement Type F504"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-FBL5N----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)
'Written This step as the date is not avaialble for the recent years.
DT_FBL5N_1000_POSTINGDATE= "01.01."& CSTR(Year(Date)-5)
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDate(DT_FBL5N_1000_POSTINGDATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FBL5N_1000_POSTING_DATE),False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot

Call ClickLabel("Typ", "0", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Type","%%DYN001-LOW","",DT_FBL5N_0841_SEARCH_TERM_OCC2,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButton("Find",True)
Call TakeScreenShot
Call SetTextbox("Find","GD_SEARCHSTR","",DT_FBL5N_0841_SEARCH_TERM,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call TakeScreenShot

Call ClickButton("Find",True)
Call TakeScreenShot
Call SetTextbox("Find","GD_SEARCHSTR","",DT_FBL5N_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call TakeScreenShot

Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot

Call ClickLabel("Text", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Text","%%DYN002-LOW","",DT_FBL5N_1105_TEXT,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

CAll SetHorizontalScrollBar(110, FAlse)
Call TakeScreenShot

Call ClickLabel("Profit Ctr", "0", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Profit Center","%%DYN003-LOW","",DT_FBL5N_1105_PROFIT_CENTER,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot


Call CLickButton("Display Document   \(Shift\+F2\)",False)
Call ClickButton("Call Up Document Overview   \(F9\)",False)

Call GetGridContentByTitle("", 0, "Amount", 1, "DT_AMOUNT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_AMOUNT_OUTPUT",DT_AMOUNT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("", 0, "Amount", 2, "DT_AMOUNT_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_AMOUNT_OCC1_OUTPUT",DT_AMOUNT_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type","BKPF-BLART", 0, DT_FBL5N_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call CLickButton("Cancel   \(F12\)",True)
Call LogOff'
Call FinalStatus()
