

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.09 Store Cash Management_Credit cards manual proceV4_TASE
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

gstrTestCaseName = "Test_09.07.01.02.09 Store Cash Management_Credit cards manual proceV4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode F-21----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Period","BKPF-MONAT","",DT_F21_0100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",DT_F21_0100_DOCUMENT_DATE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F21_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F21_0100_DOCHEADER_TEXT,False)
Call TakeScreenShot()

Call FocusTextBox("Account","RF05A-NEWKO",False)
Call PressEnter()
Call TakeScreenShot()
Call SelectCheckbox("BKPF-XMWST",0,DT_F21_0301_CALCULATE_TAX,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0301_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0301_ACCOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0301_PSTKY,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F21_0301_BUS_AREA,False)
Call TakeScreenShot()

Call FocusTextBox("Account","RF05A-NEWKO",False)
Call PressEnter() 
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0300_TEXT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F21_0300_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT,False)
Call TakeScreenShot()

Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F21_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F21_1007_BUSINESS_AREA,False)
Call TakeScreenShot()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0300_ACCOUNT,False)
Call TakeScreenShot()

Call FocusTextBox("Account","RF05A-NEWKO",False)
Call PressEnter() 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT_OCC1,False)
Call TakeScreenShot()

Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_F21_1016_PROFIT_CTRS,False)
Call TakeScreenShot()

Call FocusTextBox("Profit Ctrs","COBL-PRCTR",False)
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call SetTextbox("House Bank","BSEG-HBKID","",DT_F21_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F21_0330_BSEGHKTID,False)
Call TakeScreenShot()

Call FocusTextBox("/","BSEG-HKTID",False)
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0300_TEXT_OCC1,False)
Call TakeScreenShot()

Call FocusTextBox("Text","BSEG-SGTXT",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_F21_0100_CHECK_TEXT_OF_STATUSBAR)

''''Tcode FBL5N ''''''''''''''

Call SetTcode(DT_F21_0100_OKCD)
Call PressEnter()
Call SelectRadioButton("X_AISEL","All Items",False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F21_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot()

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_F21_1106_DOCUMENT_NUMBER,True)
Call ClickBUtton("Copy   \(F8\)",True)
Call ClickBUtton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC2,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC3,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC4,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)


Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "Text", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "BUDAT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "GSBER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)

''''Tcode FAGLL03 ''''''''''''''

Call SetTcode(DT_F21_0500_OKCD)
Call PressEnter()
Call SelectRadioButton("X_AISEL","All Items",False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_F21_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_F21_1000_TO),False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F21_1000_GL_ACCOUNT,False)
Call TakeScreenShot()

Call FocusTextBox("to","SO_BUDAT-HIGH",False)
Call ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC4)

Call SelectColumnGuiGrid("", 0, "Document Number", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F21_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F21_0841_SEARCH_TERM_OCC6,True)
Call SetComboByKey("Search Direction",DT_F21_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
'Call VerifyGridCellContent("", 1, "GSBER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
'Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER_OCC1)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)
Call VerifyGridCellContent("", 1, "BUDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT_OCC1))


''''''''''''''  LogOff from Application ''''''''''''''''

Call LogOff()
Call FinalStatus ()

