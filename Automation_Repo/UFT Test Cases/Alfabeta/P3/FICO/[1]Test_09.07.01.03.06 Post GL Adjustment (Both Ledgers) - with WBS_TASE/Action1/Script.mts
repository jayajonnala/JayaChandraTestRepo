		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.03.06 Post GL Adjustment (Both Ledgers) - with WBS
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

gstrTestCaseName = "Test_09.07.01.03.06 Post GL Adjustment (Both Ledgers) - with WBS"
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

'''''''''--------TransactionCode-F-02----------''''
'''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubleDigit(Cstr(MOnth(DT_F02_0100_PERIOD))),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()


Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("WBS element","COBL-PS_POSID","",DT_F02_1007_WBS_ELEMENT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER_OCC1,False)

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot

Call VerifyTextBoxContent("C","RF05A-AZSAL", 0, DT_F02_0700_CHECK_TEXT_OF_C, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''--------TransactionCode-fb03----------''''
Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F02_0100_DOCUMENT_NUMBER,False) 
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F02_0100_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",Year(DT_F02_0100_FISCAL_YEAR),False) 
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F02_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F02_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)
	
Call VerifyGridCellContent("", 1, "PROJK_EXT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROJK_EXT)
Call VerifyGridCellContent("", 2, "PROJK_EXT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PROJK_EXT)

Call LogOff()
Call FinalStatus()

