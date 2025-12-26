		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.04.03 Mass Reverse GL Documents
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

gstrTestCaseName = "Test_09.07.01.04.03 Mass Reverse GL Documents"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''''''''--------TransactionCode-F.80----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_F80_1000_COMPANY_CODE,False)
Call SetTextbox("Reason for reversal","STOGRD","","",False)
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",Year(DT_F80_1000_FISCAL_YEAR),False)
Call SetTextbox("Reason for reversal","STOGRD","",DT_F80_1000_REASON_FOR_REVERSAL,False)
Call SetTextbox("Posting Date","STODAT","",ConvertDate(DT_F80_1000_POSTING_DATE),False)
Call SetTextbox("Posting period","MONAT","",ConvertDoubledigit(Cstr(Month(DT_F80_1000_POSTING_PERIOD))),False)

Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_0X,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_1X,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call VerifyifGuiLabelExistsByRelativeid(Lcase(DT_F80_0120_CHECK_TEXT_OF_ALL_DOCUMENTS_CAN_BE_REVERSED), "wnd\[0\]/usr/lbl\[0,18\]")

Call Clickbutton("Back   \(F3\)",False)
Call SelectCheckbox("TESTLAUF",0, "OFF", False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call GetLabelContentByRefLabel("Document Number", -112, -64, "DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_800133959_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_800133959_OUTPUT",DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_800133959)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetLabelContentByRefLabel("Document Number", -112, -96, "DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_800133960_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_800133960_OUTPUT",DT_F80_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_800133960)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyifGuiLabelExistsByRelativeid(Lcase(DT_F80_0120_CHECK_TEXT_OF_ALL_DOCUMENTS_WERE_REVERSED), "wnd\[0\]/usr/lbl\[0,21\]")

'''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_F80_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot


Call ClickButton("%_SD_SAKNR_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call SelectRadioButton("X_AISEL","All items", False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#4;#1")

Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_0_OCC2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_1_OCC2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_F80_3010_TABLECELL_SINGLE_VALUE_3,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 7, "ICO_AUGP", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_ICO_AUGP)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 3, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 4, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)
Call VerifyGridCellContent("", 7, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BELNR)
Call VerifyGridCellContent("", 8, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BELNR)
Call VerifyGridCellContent("", 9, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_BELNR)
Call VerifyGridCellContent("", 10, "BELNR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BELNR)


Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 7, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_BSCHL)
Call VerifyGridCellContent("", 8, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_BSCHL)
Call VerifyGridCellContent("", 9, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BSCHL)
Call VerifyGridCellContent("", 10, "BSCHL", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BSCHL)


Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
Call VerifyGridCellContent("", 7, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_DMSHB)
Call VerifyGridCellContent("", 8, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_DMSHB)
Call VerifyGridCellContent("", 9, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_DMSHB)
Call VerifyGridCellContent("", 10, "DMSHB", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_DMSHB)

Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)

Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 3, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
Call VerifyGridCellContent("", 4, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PRCTR)
Call VerifyGridCellContent("", 7, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_PRCTR)
Call VerifyGridCellContent("", 8, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_PRCTR)
Call VerifyGridCellContent("", 9, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_PRCTR)
Call VerifyGridCellContent("", 10, "PRCTR", 0, DT_F80_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_PRCTR)

Call LogOff()
Call FinalStatus()
