		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.16.04 Card + Payment Deficit_Movement Type F408

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

gstrTestCaseName = "Test_09.06.05.16.04 Card + Payment Deficit_Movement Type F408_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)
DT_FAGLL03_1000_POSTING_DATE= "01.01."& CSTR(Year(Date)-5)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FAGLL03_1000_TO),False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",false)
Call ActivateNodeGuiTree("0","#4;#8")
Call SetTextBox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot

Call SetHorizontalScrollBar(75,False)
Call TakeScreenshot

Call ClickLabel("Text", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Text","%%DYN001-LOW","",DT_FAGLL03_1105_TEXT,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call CLickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Call TakeScreenshot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenshot

Call GetGridContentByTitle("", 0, "Amount", 1, "DT_FAGLL03_0500_GRIDCELL_0_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Amount", 2, "DT_FAGLL03_0500_GRIDCELL_0_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC1_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FAGLL03_0500_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FAGLL03_0500_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)


Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("", 2, "HKONT", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)


Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type","BKPF-BLART", 0, DT_FAGLL03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call CLickButton("Cancel   \(F12\)",True)
Call LogOff'
Call FinalStatus()
