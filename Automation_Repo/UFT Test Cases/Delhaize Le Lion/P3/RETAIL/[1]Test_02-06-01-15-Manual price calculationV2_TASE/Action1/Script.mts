'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_06-09-09-Manual Po in SAP
'.................Test Scenario: AT_04-06-02-06-04-Vendor returns-DC Drinks-No PO-Tied Empties
'.................TCode: ME21N,ME23N
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "TC3_Test_02-06-01-15-Manual V2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//----------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------//

'Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)'.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'-------------------------------VKP5----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextBox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)
Call SetTextBox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call SetTextBox("Sales Organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("Price List","S_PLTYP-LOW","",DT_VKP5_1000_PRICE_LIST,False)
Call TakeScreenShot()
Call SetTextBox("Purchase Price Determ\. Seq\.","P_EKERV","",DT_VKP5_1000_PURCHASE_PRICE_DETERM_SEQ,False)
Call ClickButton("Execute   \(F8\)",False)
Call PressEnter()
Call TakeScreenShot()
Call SelectCheckboxNoLabel(0,"ON",False)
Call TakeScreenShot()
Call ClickButton("Sales   \(F6\)",False)
Call TakeScreenShot()

Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",11,"","","DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_10_OUTPUT",false)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",12,"","","DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_172_OUTPUT",false)

Call VerifyTableCellContent(15,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CNTY_13)
Call VerifyTableCellContent(15,"Description","SAPLV69ATCTRL_KONDITIONEN",LCase(DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_NAME_13))
Call VerifyTableCellContent(15,"Amount","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_13)
Call VerifyTableCellContent(15,"Crcy","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CRCY_13)
'Call VerifyTableCellContent(15,"Condition value","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CONDITION_VALUE_13)

Call VerifyTableCellContent(16,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CNTY_14)
Call VerifyTableCellContent(16,"Description","SAPLV69ATCTRL_KONDITIONEN",UCase(DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_NAME_14))
Call VerifyTableCellContent(16,"Amount","SAPLV69ATCTRL_KONDITIONEN","5,000")
Call VerifyTableCellContent(16,"Crcy","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CRCY_14)
Call VerifyTableCellContent(16,"Condition value","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CONDITION_VALUE_14)
Call VerifyTableCellContent(15,"Curr.","SAPLV69ATCTRL_KONDITIONEN",DT_VKP5_6201_CHECK_TEXT_OF_TABLECELL_CURR_14)

Call VerifyTextBoxContent("Material","KOMP-MATNR","",DT_VKP5_9000_CHECK_TEXT_OF_ARTICLE,False)

Call SelectRowGuiTableByRow("SAPLV69ATCTRL_KONDITIONEN",15,False)
Call SelectRowGuiTableByRow("SAPLV69ATCTRL_KONDITIONEN",16,False)
Call ClickButton("Back   \(F3\)",False)
Call VerifyifGuiLabelExistsByRelativeid(DT_VKP5_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[17,0\]")

' ''GetLabelContentByRefLabel(refLabelContent, xDifferenceValue, yDifferenceValue, dataTableColumnName, blnIsItPopup)
'''Call GetLabelContentByRefLabel("7167399",0,-16,"DT_VKP5_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",false)
'Call GetLabelContentByRefLabel(DT_VKP5_1000_ARTICLE,0,-16,"DT_VKP5_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",false)
Call GetLabelContentByRefLabel(DT_VKP5_1000_ARTICLE,0,-18,"DT_VKP5_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",false)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("item1","DT_VKP5_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_VKP5_1000_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_VKP5_1000_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusbar(DT_VKP5_1000_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call LogOff()
Call FInalStatus()

