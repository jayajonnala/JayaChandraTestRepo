		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.09.02 Change GL Documents
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

gstrTestCaseName = "Test_09.07.01.09.02 Change GL Documents"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-FB03 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", YEar(DT_FB03_0100_FISCAL_YEAR), False)
Call PressEnter()     
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB03_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB03_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "Text", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 2, "Text", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_FB50L_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)

'''''--------TransactionCode-FB09 ----------''''

Call SetTcode(DT_FB03_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0102_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0102_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", YEar(DT_FB03_0102_FISCAL_YEAR), False)
'Line item is changed to item
Call SetTextboxNoLabel("RF05L-BUZEI","",DT_FB03_0102_LINE_ITEM,False)
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_FB03_0300_TEXT,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call VerifyStatusBar(LCase(DT_FB03_0102_CHECK_TEXT_OF_STATUSBAR))
'Line item is changed to item
Call SetTextboxNoLabel("RF05L-BUZEI","",DT_FB03_0102_LINE_ITEM_OCC1,False)
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_FB03_0300_TEXT_OCC1,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call VerifyStatusBar(LCase(DT_FB03_0102_CHECK_TEXT_OF_STATUSBAR_OCC1))


'''''--------TransactionCode-FB03 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", YEar(DT_FB03_0100_FISCAL_YEAR_OCC1), False)
Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "Text", 0, LCase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1))
Call VerifyGridCellContent("", 2, "Text", 0, LCase(DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT_OCC1))

Call VerifyGridCellContent("", 1, "Account", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "Account", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)

Call LogOff()
Call FinalStatus()


