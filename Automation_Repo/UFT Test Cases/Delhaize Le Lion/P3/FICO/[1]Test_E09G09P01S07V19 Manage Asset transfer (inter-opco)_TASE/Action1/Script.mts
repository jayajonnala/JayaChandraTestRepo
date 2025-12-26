'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S07V19 Manage Asset transfer (inter-opco)  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S07V19 Man(inter-opco)"
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

''''''--------TransactionCode-ABT1N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABT1N_0300_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("RAIFP2-XERBW","Rev. from NBV", False)
Call TakeScreenShot

Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABT1N_0300_ASSET,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABT1N_0300_ASSET_OCC1,False)

Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABT1N_0200_DOCUMENT_DATE),False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABT1N_0202_ASSET_VALUE_DATE),False)
Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABT1N_0206_TEXT,False)
Call SetTextboxNoLabel("RAIFP2-AFABE", "", DT_ABT1N_0410_RAIFP2AFABE, False)
Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_ABT1N_0330_EXISTING_ASSET,False)
Call SetTextbox("Existing asset","RAIFP3-ANLN2","",DT_ABT1N_0330_EXISTING_ASSET_OCC1,False)
Call SetTextbox("Company Code","RAIFP3-BUKRS","",DT_ABT1N_0321_COMPANY_CODE,False)

Call ClickButton("Simulate   \(F9\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetLabelContentByRefLabel("Message text","0","-32", "DT_OP_ABT1N_0120_CHECK_TEXT_OF_ASSET_TRANSACTION_POSTED_WITH_DOCUMENT_NO_BE20_0100000550", True)
Call GetLabelContentByRefLabel("Message text","0","-48", "DT_OP_ABT1N_0120_CHECK_TEXT_OF_ASSET_TRANSACTION_POSTED_WITH_DOCUMENT_NO_BE10_0100015636", True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyifGuiLabelExists_ByIndex("Asset transaction posted with document no. BE20 "&DT_DOC1,0)
Call VerifyifGuiLabelExists_ByIndex("Asset transaction posted with document no. BE10 "&DT_DOC2,0)

Call ClickButton("Continue   \(Enter\)",True)

''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_ABT1N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)


Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","","")
Call TakeScreenShot
Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change","","")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Year-end","","")
Call TakeScreenShot
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABT1N_0201_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABT1N_0201_ASSET,False)
Call TakeScreenShot
Call PressEnter() 
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_VERAENDERUNG_OCC3)
'Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Year-end","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OCC3)
Call TakeScreenShot
Call SelectNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","",DT_ABT1N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OCC2)
Call TakeScreenShot
Call DoubleClickGuiGridCell("Transactions", "", 1, "Amount posted", False)

''''''--------TransactionCode-FB03----------''''
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Profit Center", "", DT_ABT1N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

