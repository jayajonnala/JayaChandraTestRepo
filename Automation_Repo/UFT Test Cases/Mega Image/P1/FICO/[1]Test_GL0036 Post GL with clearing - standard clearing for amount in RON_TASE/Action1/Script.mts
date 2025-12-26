
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0036 Post GL with clearing - standard clearing for amount in RON_TASE
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


gstrTestCaseName = "Test_GL0036 Post GL with clearing - standard clearing for amount in RON_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''''--------TransactionCode-F-02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F02_0100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)

Call TakeScreenShot
''Call PressEnter()
''Call TakeScreenShot
''Call PressEnter()     
''Call TakeScreenShot
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=31","attachedtext:=Reference","name:=BKPF-XBLNR","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0300_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetSpecialTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC1,False)


Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOCUMENT_NO_OUTPUT&" was posted in company code RO02")
'
'''''''''''--------TransactionCode-F-04----------''''

Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0122_POSTING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F02_0122_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0122_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0122_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0122_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0122_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0122_DOCHEADER_TEXT,False)

Call TakeScreenShot
''Call PressEnter()
''Call TakeScreenShot
''Call PressEnter()     
''Call TakeScreenShot
While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=31","attachedtext:=Reference","name:=BKPF-XBLNR","Index:=0").Exist(5)=True 
Call PressEnter()  
Wend

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA_OCC1,False)
Call TakeScreenShot
Call ClickBUtton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SelectCheckbox("BSEG-XNEGP",0,"ON",False)
Call TakeScreenShot
Call ClickBUtton("Choose open items   \(F6\)",False)
Call TakeScreenShot
Call SelectradioBUtton("RF05A-XPOS1","Document Number",False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F02_0710_ACCOUNT,False)
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SetTextbox("From","RF05A-SEL01",0,DT_DOCUMENT_NO_OUTPUT,False)
Call TakeScreenShot
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)

Call SelectMenuBar("Document;Simulate")
Call PressEnter()

Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F02_0700_DOCHEADER_TEXT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_DOCUMENT_NO2_OUTPUT")


'''----------------------Tcode FBL3N----------------------------

Call SetTcode(DT_F02_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F02_1000_GL_ACCOUNT,False)
Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)

Call ActivateNodeGuiTree(0, "Document;Document Number")
Wait 5
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)
Wait 2
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_DOCUMENT_NO_OUTPUT, true)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_DOCUMENT_NO2_OUTPUT, true)
Call ClickButton("Copy   \(F8\)",True)
Wait 2
Call SelectRadioBUtton("X_AISEL","All items",False)

Call ClickBUtton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")

Call ClickLabel(DT_DOCUMENT_NO_OUTPUT, 0, False)
Call TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)

Call VerifyGridCellContent("", 1, "KOBEZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 2, "KOBEZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)

Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 1, "GSBER", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call ClickBUtton("Back   \(F3\)",False)
Wait 5
'
Call ClickLabel(DT_DOCUMENT_NO2_OUTPUT, 0, False)
Call TakeScreenShot

Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC1)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)

Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ_OCC1)

Call VerifyGridCellContent("", 1, "KOBEZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ_OCC1)
Call VerifyGridCellContent("", 2, "KOBEZ", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ_OCC1)

Call VerifyGridCellContent("", 2, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC1)
Call VerifyGridCellContent("", 1, "GSBER", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER_OCC1)
Call VerifyGridCellContent("", 1, "Profit Center", 0, DT_F02_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)

Call Logoff'
Call FinalStatus()'
