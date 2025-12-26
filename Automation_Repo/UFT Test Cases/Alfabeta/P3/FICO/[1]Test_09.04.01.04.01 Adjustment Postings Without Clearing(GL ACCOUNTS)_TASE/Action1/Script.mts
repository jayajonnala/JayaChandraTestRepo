

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.04.01 Adjustment Postings Without Clearing(GL ACCOUNTS)
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

gstrTestCaseName = "Test_09.04.01.04.01 Adjustment Postings Without Clearing(GL ACCOUNTS)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.04.01 Adjustment Postings Without Clearing(GL ACCOUNTS).xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''--------TransactionCode-F-02 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubledigit(CSTR(Month(DT_F02_0100_PERIOD))),False)
Call SetTextbox("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_F02_0100_POSTING_DATE), False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_F02_0100_DOCUMENT_DATE), False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot


Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0302_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0302_TEXT,False)
Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_F02_0302_GL_ACC,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_F02_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0302_TAX_CODE,False)
Call TakeScreenShot
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0302_ACCOUNT,False)
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F02_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F02_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F02_1007_BUSINESS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet("DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot


'''''--------TransactionCode-FAGLL03 ----------''''

Call SetTcode(DT_F02_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL","All items", False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_F02_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_F02_1000_TO),False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F02_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_F02_1000_COMPANY_CODE,False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",false)
Call ActivateNodeGuiTree("0","#4;#1")
Call SetTextBox("Document Number","%%DYN001-LOW","",DT_F02_0100_DOCUMENT_NUMBER,False)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F02_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_F02_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F02_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_F02_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)


Call ClickButtonIfExist("Transfer   \(Enter\)",True)


Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
'Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "Tax Code", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Text", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

'''' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
''''Call GetGridContent("","","Assignment",1,"Document Number",DT_F02_0100_DOCUMENT_NUMBER,"DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT")

Call GetGridContent("","","G/L Account",1,"Document Number",DT_F02_0100_DOCUMENT_NUMBER,"DT_F02_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OUTPUT")

Call LogOff'
Call FinalStatus()





''''Call ClickLabel("DocumentNo", "1", False)
'''Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'''Call ClickButtonToolBar("&FIND",0)
'''Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
'''Call ClickButtonIfExist("Continue   \(Enter\)",True)
'''Call ClickButtonIfExist("Cancel   \(F12\)",True)
'''Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
'''Call ClickButtonIfExist("Define Filter Values",True)
'''
'''Call ClickButtonIfExist("B_SEARCH",True)
'''Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
'''Call ClickButtonIfExist("Continue   \(Enter\)",True)
'''Call ClickButtonIfExist("Show sel\. fields \(CTRL\+F3\)",True)
'''Call ClickButtonIfExist("Copy   \(Enter\)",True)
''''
'''Call SetTextbox("Document Number G/L","%%DYN001-LOW","",DT_F02_0100_DOCUMENT_NUMBER,True)
''''''Call SetTextbox("Document Number","%%DYN001-LOW","",DT_F02_0100_DOCUMENT_NUMBER,True)
'''Call ClickButton("Execute   \(Enter\)",True)
'''Call TakeScreenshot
'''
'Call CLickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
'Call TakeScreenshot
'Call ClickButton("Call Up Document Overview   \(F9\)",False)
'Call TakeScreenshot
'
'
