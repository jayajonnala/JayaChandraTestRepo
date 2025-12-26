		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_2_09.12.02.03.03 Manage Account Receivable Monthly Closing
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

gstrTestCaseName = "Test_PRE_2_09.12.02.03.03 Manage Account Receivable Monthly Closing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-SM35----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Process session   \(F8\)",False)
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call ClickButton("Process   \(Enter\)",True)
Call ClickButton("Go back to batch input session overview   \(Enter\)",True)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Analyze session   \(F2\)",False)


DT_TABNAME = " Log created on "&ConvertDate(DT_SM35_0100_LOG_CREATED_ON_DATE)
Call SelectTab("TAB_DYNPRO",DT_TABNAME,False)
Call TakeScreenShot

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message","3", "", "", "DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2_OUTPUT",DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message","5", "", "", "DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4_OUTPUT",DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo) 

Call VerifyTableCellContent("3", "Message", "RSBDC_ANALYSETC_PROTOCOL", Lcase(DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2_OCC1))
Call VerifyTableCellContent("5", "Message", "RSBDC_ANALYSETC_PROTOCOL", Lcase(DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4_OCC1))

''''''----------------TransactionCode-FBL5N-----------------------------------------------

Call SetTcode(DT_SM35_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_SM35_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_SM35_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_CLSEL","Cleared items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_SM35_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_SM35_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_SM35_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_SM35_3010_TABLECELL_SINGLE_VALUE_3,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "BELNR", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 3, "BELNR", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call VerifyGridCellContent("", 4, "BELNR", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BELNR)

Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)

'''''''----------------TransactionCode-FAGLL03-----------------------------------------------

Call SetTcode(DT_SM35_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_SM35_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_SM35_1000_COMPANY_CODE_OCC1,False)
Call SelectRadioButton("X_AISEL","All Items", False)
Call TakeScreenShot
Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#4;#1")
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_SM35_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_SM35_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_SM35_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_SM35_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "DMSHB", 0,DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0,DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0,DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 1, "UMSKZ", 0,DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 2, "UMSKZ", 0,DT_SM35_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
Call Logoff()
Call FinalStatus()


