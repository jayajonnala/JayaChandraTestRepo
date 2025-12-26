		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.09.02 Asset Acquisition without supplier
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


gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.11.01.09.02 Asset Acquisition without supplier_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'''''--------TransactionCode-AS01----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call PressEnter()
Call TakeScreenShot()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Quantity","ANLA-MEINS","",DT_AS01_1140_QUANTITY,False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Origin",False)
Call TakeScreenShot()
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","Allocations",False)
Call TakeScreenShot()
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","Time-dependent",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","General",False)
Call TakeScreenShot()
Call Clickbutton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-f-90 ----------''''
Call SetTcode(DT_AS01_0105_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_AS01_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS01_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_AS01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_AS01_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_AS01_1007_COST_CENTER,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0300_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0300_TTYPE,False)
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0305_TAX_CODE,False)
Call SetTextbox("Quantity","BSEG-MENGE","",DT_AS01_0305_QUANTITY,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0300_PSTKY_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call Clickbutton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call PressEnter()
Call TAkescreenshot()
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectMenuBar("Document;Display")
Call TakeScreenShot()

Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("", 2, "MWSKZ", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)

''''''--------TransactionCode-nas03 ----------''''

Call SetTcode(DT_AS01_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 1, "BWATXT", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWATXT)
Call VerifyGridCellContent("Transactions", 1, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS)

Call ActivateNodeGuiTree("0","Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC1)
Call VerifyGridCellContent("Transactions", 1, "BWATXT", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWATXT_OCC1)
Call VerifyGridCellContent("Transactions", 1, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS_OCC1)

Call ActivateNodeGuiTree("0","Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call VerifyGridCellContentbyName("shell", 1, "Value", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BEZEICHNUNG)
Call VerifyGridCellContentbyName("shell", 1, "Change", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContentbyName("shell", 1, "Year-end",0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContentbyName("shell", 1, "Currency", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS_OCC2)

Call VerifyGridCellContentbyName("shell", 4, "BEZEICHNUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BEZEICHNUNG)
Call VerifyGridCellContentbyName("shell", 4, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContentbyName("shell", 4, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call VerifyGridCellContentbyName("shell", 11, "BEZEICHNUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_BEZEICHNUNG)
Call VerifyGridCellContentbyName("shell", 11, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_VERAENDERUNG)
Call VerifyGridCellContentbyName("shell", 11, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE)
Call VerifyGridCellContentbyName("shell", 11, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_WAERS)

Call ActivateNodeGuiTree("0","Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

Call VerifyGridCellContentbyName("shell", 1, "BEZEICHNUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BEZEICHNUNG_OCC1)
Call VerifyGridCellContentbyName("shell", 1, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContentbyName("shell", 1, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC1)
Call VerifyGridCellContentbyName("shell", 1, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS_OCC3)

Call VerifyGridCellContentbyName("shell", 4, "BEZEICHNUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BEZEICHNUNG_OCC1)
Call VerifyGridCellContentbyName("shell", 4, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContentbyName("shell", 4, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC1)
Call VerifyGridCellContentbyName("shell", 4, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_WAERS)

Call VerifyGridCellContentbyName("shell", 11, "BEZEICHNUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_BEZEICHNUNG_OCC1)
Call VerifyGridCellContentbyName("shell", 11, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_VERAENDERUNG_OCC1)
Call VerifyGridCellContentbyName("shell", 11, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OCC1)
Call VerifyGridCellContentbyName("shell", 11, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_WAERS_OCC1)

Call LogOff'
Call FinalStatus()


