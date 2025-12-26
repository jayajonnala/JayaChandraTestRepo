		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.06.01 Western Union Out_Movement Type F502

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

gstrTestCaseName = "Test_09.06.05.06.01 Western Union Out_Movement Type F502"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-FBL5N----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDate(DT_FBL5N_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_POSTING_DATE_TO),False)

Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_SEARCH_TEXT,True)
Call SetComboByKey("Search Direction",DT_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_SEARCH_TEXT_OCC1,True)
Call SetComboByKey("Search Direction",DT_SEARCH_DIRCT_OCC1)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call SelectColumnGuiGrid("", "0", "Profit Center", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",false)
Call TakeScreenShot
Call SetTextboxNoName("Profit Center","",DT_FBL5N_1105_PROFIT_CENTER,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
' SetTextboxNoName(textboxAttachedText, textboxIndex, textboxValue, blnIsItPopup)



Call SelectColumnGuiGrid("", "0", "Text", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",false)
Call TakeScreenShot
Call SetTextboxNoName("Text","",DT_FBL5N_1105_TEXT,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call DoubleClickGuiGridCell("", "0", "1", "Document type", False)
Call TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)

Call GetGridContentByTitle("", 0, "Amount", 1, "DT_AMOUNT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_AMOUNT_OUTPUT",DT_AMOUNT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("", 0, "Amount", 2, "DT_AMOUNT_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_AMOUNT_OCC1_OUTPUT",DT_AMOUNT_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)
Call LogOff'
Call FinalStatus()
