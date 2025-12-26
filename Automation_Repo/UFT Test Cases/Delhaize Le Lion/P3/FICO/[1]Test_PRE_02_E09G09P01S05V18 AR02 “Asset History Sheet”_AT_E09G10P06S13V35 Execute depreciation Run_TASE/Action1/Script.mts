'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_02_E09G09P01S05V18 AR02 “Asset History Sheet”  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_02_E09G09P01S05V18 As H"
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

'''''''--------TransactionCode-F-02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F02_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F02_0100_TTYPE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F02_0100_PERIOD,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0305_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0305_TEXT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0305_TAX_CODE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0305_ACCOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0305_PSTKY,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call TakeScreenShot
Call ClickButton("All Acct Assignmts", False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_0002_BUSINESS_AREA,True)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F02_0002_PROFIT_CENTER,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)", True)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)  
Call TakeScreenShot
Call GetStatusBar("item1","DT_OP_DOC_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_F02_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

''''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Asset","ANLA-ANLN1","",DT_F02_0201_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_F02_0201_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN2","",DT_F02_0201_ASSET_OCC1,False)
Call TakeScreenShot
Call SetTextbox("Fiscal year","EDIT_JAHRE","",DT_F02_0202_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call TakeScreenShot

Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)
Call TakeScreenShot

Call SelectNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);11 Difference Local - IFRS Depreciation (6-1)")
'Call VerifyGridCellContent("Planned values Difference Local - IFRS Depreciation (6-1)",6,"Change","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_VERAENDERUNG)
'Call VerifyGridCellContent("Planned values Difference Local - IFRS Depreciation (6-1)",6,"Year-end","",DT_F02_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE)
'Call VerifyGridCellContent("Planned values Difference Local - IFRS Depreciation (6-1)",1,"Change",""," ")
'Call VerifyGridCellContent("Planned values Difference Local - IFRS Depreciation (6-1)",1,"Year-end",""," ")
Call VerifyGridCellContent("Planned values Difference Local - IFRS Depreciation (6-1)",1,"Change","","")
Call VerifyGridCellContent("Planned values Difference Local - IFRS Depreciation (6-1)",1,"Year-end","","")
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

