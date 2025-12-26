
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0006_Manage_AR_Documents_Manage_Manual_Customer_Invoicing_Cred 
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


gstrTestCaseName = "Test_AR0006_Manage_AR_Documents_Manage_Manual_Customer_Invoicing_Cred"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0006_Manage_AR_Documents_Manage_Manual_Customer_Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''''--------TransactionCode-F-22----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F22_100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F22_100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F22_100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F22_100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F22_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F22_100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F22_100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_100_ACCOUNT,False)
Call TakeScreenShot
'Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_301_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F22_301_ACCOUNT,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_301_ASSIGNMENT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_300_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F22_301_PSTKY,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Text","BSEG-SGTXT","",DT_F22_300_TEXT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F22_300_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F22_300_AMOUNT,False)
Call TakeScreenShot
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1", "DT_DOC_NO_OUTPUT")
Call TakeScreenShot
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" was posted in company code RO02")
Call SelectMenuBar("Document;Display")
Call SelectColumnGuiGrid("", 0, "G/L Account", False)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F22_SEARCHTEXT,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F22_SEARCHTEXT,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButtonToolBar("DTC_UPPOS1", 0)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call DoubleClickGuiGridCell("",0, 1, "Account", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC2, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0, 2, "Account", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC2)
'Call VerifyGridCellContent("", 2, "HKONT", 0, DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F22_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call GetTextboxValue("BKPF-XBLNR", "", "DT_GET_TEXT_OF_REFERENCE_OUTPUT", False)
Call TakeScreenShot
Call ClickButton("Display Document Header   \(F5\)",False)
Call GetTextboxValue("BKPF-KURSF", "", "DT_F22_1710_CHECK_TEXT_OF_EXCHANGE_RATE_OUTPUT", True)
Call TakeScreenShot
Call ClickButton("btn\[0\]",True)

''''''''--------TransactionCode-//nob08----------''''

Call SetTcode(DT_F22_750_OKCD)
Call TakeScreenShot
Call PressEnter()     
Call PressEnter() 

Call ClickButton("VIM_POSI_PUSH",False)  ''Position button

Call SetTextboxNoLabel("SVALD-VALUE","",DT_F22_0300_EXCH_RATE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_F22_0300_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_F22_0300_TOCURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDate(DT_F22_0300_VALID_FROM),True)
Call ClickButton("Continue   \(Enter\)",True)
'Call VerifyTableCellContent(1, "Dir.quot.", "SAPL0SAPTCTRL_V_TCURR", DT_F22_20_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0)
Call TakeScreenShot
'
'''''''''--------TransactionCode-/nfbl5n----------''''

Call SetTcode(DT_F22_20_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F22_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F22_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call SetTextbox("Layout","PA_VARI","",DT_LAYOUT,False)
Call SelectRadioButton("X_AISEL", "All Items", False)
Call ClickButton("Execute   \(F8\)", False)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot
Call ClickLabel("DocumentNo", "", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)", False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_DOC_NO_OUTPUT,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Change layout...   \(Ctrl\+F8\)", False)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",61, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",62, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST",70, True)
Call ClickButton("Show Sel. Fields \(Ctrl\+F3\)", True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_DOC_NO_OUTPUT,0)

Call ClickButton("Last column   \(Ctrl\+F4\)", False)
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_F22_120_CHECK_TEXT_OF_NO_NAME),0)

Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_20A1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F22_120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call TakeScreenShot
'
'Call ClickButton("Last column   \(Ctrl\+F4\)", False)
'Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_F22_120_CHECK_TEXT_OF_NO_NAME),0)
'Call TakeScreenShot


Call LogOff()
Call FinalStatus ()






'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




