'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G10P06S14V46 Impairment Testing
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G10P06S14V46 Impairment"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-ABAW----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Posting Date"," ANEK-BUDAT","",ConvertDate(DT_ABAW_0100_POSTING_DATE),False)
Call SetTextbox("Transaction Type"," ANBZ-BWASL","",DT_ABAW_0100_TRANSACTION_TYPE,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_ABAW_0100_DOCUMENT_DATE),False)
Call SetTextbox("Sub-number","ANBZ-ANLN2","",DT_ABAW_0100_SUBNUMBER,False)
Call SetTextbox("Asset","ANBZ-ANLN1","",DT_ABAW_0100_ASSET,False)
Call SetTextbox("Company Code","ANBZ-BUKRS","",DT_ABAW_0100_COMPANY_CODE,False)
Call SetTextbox("Posting period","ANBZ-PERID","",DT_ABAW_0100_POSTING_PERIOD,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMA01BTCTRL_0200",2, True)
'Call SelectRowGuiTableByRow("SAPMA01BTCTRL_0200",3, True)
'Call SelectRowGuiTableByRow("SAPMA01BTCTRL_0200",4, True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot      
Call SetTextbox("Amount posted","ANBZ-DMBTR","",DT_ABAW_0160_AMOUNT_POSTED,False)
Call SetTextbox("Text","ANEK-SGTXT","",DT_ABAW_0160_TEXT,False)
Call SetTextbox("Asset Val. Date","ANBZ-BZDAT","",ConvertDate(DT_ABAW_0160_ASSET_VAL_DATE),False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Line Items   \(Shift\+F1\)", False)
Call VerifyTableCellContent(1, "Amount posted","SAPMA01BTCTRL_ANBTR", DT_ABAW_0250_CHECK_TEXT_OF_TABLECELL_AMOUNT_POSTED_0)
Call VerifyTableCellContent(2, "Amount posted","SAPMA01BTCTRL_ANBTR", DT_ABAW_0250_CHECK_TEXT_OF_TABLECELL_AMOUNT_POSTED_0)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)  
Call TakeScreenShot
Call GetStatusBar("item1","DT_OP_AA_DOC")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_ABAW_0100_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Back   \(F3\)", False)


''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_ABAW_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",3,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",3,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",10,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",10,"Year-end","","")
Call SelectTab("IDC_TABSTRIP", "Posted values", False)
Call TakeScreenShot
Call VerifyGridCellContent("Depreciation posted/planned",1,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NAFAZ)
Call VerifyGridCellContent("Depreciation posted/planned",2,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFAZ)
Call VerifyGridCellContent("Depreciation posted/planned",3,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFAZ)
Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Depreciation posted/planned",1,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NAFAZ)
Call VerifyGridCellContent("Depreciation posted/planned",2,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFAZ)
Call VerifyGridCellContent("Depreciation posted/planned",3,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFAZ)
Call VerifyGridCellContent("Depreciation posted/planned",4,"Spec.dep.to be pstd.","",DT_ABAW_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFAZ)
Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR", "", DT_ABAW_0304_CHECK_TEXT_OF_USEFUL_LIFE, False)
Call VerifyTextBoxContent("/","AW01_DEP_PAR-NDPER", "", DT_ABAW_0304_CHECK_TEXT_OF_EXP_USEFL_LIFE, False)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

