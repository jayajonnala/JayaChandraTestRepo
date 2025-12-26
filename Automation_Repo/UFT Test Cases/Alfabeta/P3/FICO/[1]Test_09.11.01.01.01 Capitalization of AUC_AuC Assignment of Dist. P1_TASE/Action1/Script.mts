		


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.01.01 Capitalization of AUC_AuC Assignment of Dist. P1
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

gstrTestCaseName = "Test_09.11.01.01.01 Capitalization of AUC_AuC Assignment of Dist. P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-AIAB----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Company Code","AICOM-BUKRS","",DT_AIAB_0110_COMPANY_CODE,False)
Call SetTextbox("Asset","AICOM-ANLN1","",(DT_AIAB_0110_ASSET),False)
Call SetTextbox("Sub-number","AICOM-ANLN2","",DT_AIAB_0110_SUBNUMBER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

CAll SelectRowGuiGridbyRowNo("", "", 1, False)
Call TakeScreenShot
Call ClickButton("Enter distribution rules   \(Shift\+F6\)",False)
Call TakeScreenShot

Call SetTableData("SAPLKOBSTC_RULES", "Cat", 1, "", "", DT_AIAB_0130_TABLECELL_CAT_0, False)
Call SetTableData("SAPLKOBSTC_RULES", "Cat", 2, "", "", DT_AIAB_0130_TABLECELL_CAT_1, False)

Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 1, "", "", DT_AIAB_0130_TABLECELL_SETTLEMENT_RECEIVER_0, False)
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 2, "", "", DT_AIAB_0130_TABLECELL_SETTLEMENT_RECEIVER_1, False)

Call SetTableData("SAPLKOBSTC_RULES", "Amount", 1, "", "", DT_AIAB_0130_TABLECELL_AMOUNT_0, False)
Call SetTableData("SAPLKOBSTC_RULES", "Amount", 2, "", "", DT_AIAB_0130_TABLECELL_AMOUNT_1, False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickBUtton("Back   \(F3\)",False)
Call TakeScreenShot
CAll SelectRowGuiGridbyRowNo("", "", 1, False)
Call TakeScreenShot
Call ClickBUtton("Execute settlement   \(Shift\+F8\)",False)
Call TakeScreenShot
Call ClickBUtton("Yes",True)
Call TakeScreenShot

Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_AIAB_0100_DOCUMENT_DATE),False)
Call SetTextbox("Asset Val\. Date","ANEP-BZDAT","",ConvertDate(DT_AIAB_0100_ASSET_VAL_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_AIAB_0100_POSTING_DATE),False)
Call SetTextbox("Text","\*KOMK3-SGTXT","",DT_AIAB_0100_TEXT,False)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANBTR)
Call VerifyGridCellContent("", 2, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANBTR)
Call VerifyGridCellContent("", 3, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ANBTR)
Call VerifyGridCellContent("", 4, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ANBTR)
Call VerifyGridCellContent("", 5, "ANBTR", 0, DT_CHECK_GRIDCELL_4)
Call VerifyGridCellContent("", 6, "ANBTR", 0, DT_CHECK_GRIDCELL_5)
Call ClickBUtton("Back   \(F3\)",False)
Call TakeScreenShot

Call SelectCheckbox("LKO74-TESTLAUF", 0, DT_AIAB_0100_TEST_RUN, False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_AIAB_0500_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_AIAB_0500_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_AIAB_0500_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(Lcase(DT_AIAB_0500_CHECK_TEXT_OF_STATUSBAR))

Call LogOff'
Call FinalStatus()

