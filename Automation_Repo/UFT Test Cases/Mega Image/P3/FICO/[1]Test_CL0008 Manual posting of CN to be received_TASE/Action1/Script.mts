'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0008 Manual posting of CN to be received 
'.................Author : TCS      
'................ Creation Date   
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


gstrTestCaseName = "Test_CL0008 Manual posting of CN to be received"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0008 Manual posting of CN to be received.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''''----------------------Login----------------------------
'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''----------------------Tcode F-02----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  
Call TakeScreenShot() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
     
Call SetTextbox("Period","BKPF-MONAT","",DT_F02_0100_PERIOD,False)
'Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE1),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
''''Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_AS11_0100_DOCHEADER_TEXT,False)

Call TakeScreenShot
'''Call PressEnter()
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=31","attachedtext:=Reference","name:=BKPF-XBLNR","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend
Call SelectCheckbox("BKPF-XMWST", 0, DT_F02_0300_CALCULATE_TAX, False)

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F02_0300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call TakeScreenShot
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F02_0300_SGL_IND,False)
Call PressEnter()
Call SetTextbox("Due on","BSEG-ZFBDT","",ConvertDate(DT_F02_0303_DUE_ON),False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call PressEnter()
Call VerifyTextBoxContent("C", "RF05A-AZHAB", 0, DT_F02_0700_CHECK_TEXT_OF_C, False)
Call VerifyTextBoxContent("C", "RF05A-AZSAL", 0, trim(DT_F02_0700_CHECK_TEXT_OF_C_OCC1), False)
''Call VerifyTextBoxContent("Line items", "RF05A-ANZAZ", 0, lcase(DT_F02_0700_CHECK_TEXT_OF_LINE_ITEMS), False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("item1","DT_DOC_NR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F02_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NR_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F02_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F02_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

''Call ClickButtonToolBar("&MB_VARIANT", 0)  ''not required
'''Call SelectRowGuiGridbyRowNo("", 0, 5, True)
''Call SelectRowGuiGridbyRowNo("", 0, 2, True)
''Call ClickButtonIfExist("Show Selected Fields   \(F7\)", True)
''Call ClickButtonIfExist("Transfer   \(Enter\)", True)
'''Continue   (Enter)

Call VerifyGridCellContent("", 1, "Account", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
Call VerifyGridCellContent("", 2, "Account", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)


Call LogOff()

Call FinalStatus ()

