'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-03-Create new assortment-BASA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S08V21 (intra-opco)"
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

''''''--------TransactionCode-ABUMN----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABUMN_0300_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABUMN_0300_ASSET,False)  
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABUMN_0300_ASSET_OCC1,False)  
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABUMN_0200_DOCUMENT_DATE),False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABUMN_0202_ASSET_VALUE_DATE),False)
Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_ABUMN_0320_EXISTING_ASSET,False)  
Call SetTextbox("Existing asset","RAIFP3-ANLN2","","",False)  
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_ABUMN_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_ABUMN_0100_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_ABUMN_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0201_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0300_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","","")
Call VerifyGridCellContent("Transactions",2,"Transaction type name","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWATXT_OCC1)
Call VerifyGridCellContent("Transactions",2,"Amount posted","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call TakeScreenShot
Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Year-end","","")
Call VerifyGridCellContent("Transactions",2,"Transaction type name","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWATXT_OCC1)
Call VerifyGridCellContent("Transactions",2,"Amount posted","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call TakeScreenShot
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0320_EXISTING_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0300_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call SelectNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_VERAENDERUNG_OCC2)
Call TakeScreenShot
Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_11_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Year-end","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_11_JENDE_OCC3)
Call TakeScreenShot
Call VerifyGridCellContent("Transactions",1,"Transaction type name","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWATXT)
Call VerifyGridCellContent("Transactions",1,"Amount posted","",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

