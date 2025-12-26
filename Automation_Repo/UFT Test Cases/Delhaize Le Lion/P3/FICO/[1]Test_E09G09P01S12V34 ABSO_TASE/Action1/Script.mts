'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S12V34 ABSO
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S12V34 ABSO"
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

''''''--------TransactionCode-ABSO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","ANBZ-BUKRS","",DT_ABSO_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANBZ-ANLN1","",DT_ABSO_0100_ASSET,False)  
'Call SetTextbox("Asset","ANBZ-ANLN2","",DT_ABUMN_0300_ASSET_OCC1,False)  
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_ABSO_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_ABSO_0100_POSTING_DATE),False)
Call SetTextbox("Transaction Type","ANBZ-BWASL","",DT_ABSO_0100_TRANSACTION_TYPE,False)  
Call SetTextbox("Posting period","ANBZ-PERID","",DT_ABSO_0100_POSTING_PERIOD,False)  
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Amount posted","ANBZ-DMBTR","",DT_ABSO_0110_AMOUNT_POSTED,False)  
Call SetTextbox("Text","ANEK-SGTXT","",DT_ABSO_0110_TEXT,False)
Call SetTextbox("Document type","RA01B-BLART","",DT_ABSO_0110_DOCUMENT_TYPE,False)    
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Ord. dep.","ANEA-NAFAL","",DT_ABSO_0600_ORD_DEP,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Ord. dep.","ANEA-NAFAL","",DT_ABSO_0600_ORD_DEP_OCC1,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Amount posted","ANEP-ANBTR","",DT_ABSO_0610_AMOUNT_POSTED,False)
Call PressEnter()
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_DOC")
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_ABSO_0100_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Cancel   \(F12\)",False)

''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_ABSO_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)


Call VerifyGridCellContent("Planned values IFRS APC, depreciation",9,"Change","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_8_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",9,"Year-end","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_8_JENDE)
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",9,"Change","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_8_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",9,"Year-end","",DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_8_JENDE_OCC1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
