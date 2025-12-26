		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.13 Store Cash Management_Electricity (DEH) manual pr
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

gstrTestCaseName = "Test_09.07.01.02.13 Store Cash Management_Electricity (DEH) manual pr"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-F-42----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F42_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F42_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0100_PSTKY,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F42_0100_TYPE_1,False)
Call SetTextbox("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_F42_0100_DOCUMENT_DATE), False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F42_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F42_0100_REFERENCE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_F42_0302_GL_ACC,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0302_PSTKY,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0302_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F42_0302_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0302_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
			

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0300_AMOUNT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F42_1006_BUSINESS_AREA,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F42_1006_PROFIT_CENTER,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F42_0300_TEXT,False)

Call SelectMenuBar("Document;Simulate")
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call GetStatusBar("item1","DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02" )
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F42_0100_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F42_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F42_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0100_PSTKY_OCC1,False)
Call SetTextbox("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_F42_0100_DOCUMENT_DATE_OCC1), False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F42_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F42_0100_REFERENCE_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0100_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SelectCheckbox("BKPF-XMWST", 0, DT_F42_0300_CALCULATE_TAX, False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0300_PSTKY_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0300_AMOUNT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F42_0300_TEXT_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F42_1006_BUSINESS_AREA_OCC1,False)
Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F42_1006_PROFIT_CENTER_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0300_AMOUNT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F42_0300_TAX_CODE,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F42_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F42_1007_COST_CENTER,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F42_0300_TEXT_OCC2,False)

Call SelectMenuBar("Document;Simulate")
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" was posted in company code GR02" )
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT",DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


''''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_F42_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_F42_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F42_3010_TABLECELL_SINGLE_VALUE_0,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F42_0500_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F42_0500_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)


''''''--------TransactionCode-fbl3n----------''''

Call SetTcode(DT_F42_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F42_1000_GL_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All items", False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ActivateNodeGuiTree(0, "Document;Document Number")
Call ClickButton("%_%%DYN007_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F42_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
							
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC2)
Call VerifyGridCellContent("", 1, "GSBER", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", 1, "Document Date", 0, ConvertDate(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLADT_OCC2))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC3)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC4)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER_OCC5)
Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ_OCC6)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC7)
Call logOff()
Call FinalStatus()

