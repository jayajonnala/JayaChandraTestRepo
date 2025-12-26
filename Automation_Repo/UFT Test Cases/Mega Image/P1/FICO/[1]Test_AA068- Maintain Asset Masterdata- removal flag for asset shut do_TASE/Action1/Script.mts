
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA068- Maintain Asset Masterdata- removal flag for asset shut do
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


gstrTestCaseName = "Test_AA068- Maintain Asset Masterdata- removal flag for asset shut do"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA068- Maintain Asset Masterdata- removal flag for asset shut do.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call EndDateof445PeriodByDate(DT_TODAY,"DT_END_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Asset main no. text","ANLH-ANLHTXT","",DT_AS01_1140_ASSET_MAIN_NO_TEXT,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Allocations", False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Origin", False)
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call SetTextbox("Type name","ANLA-TYPBZ","",DT_AS01_1181_TYPE_NAME,False)
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call SelectCheckbox("ANLZ-XSTIL", 0, DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN, False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_DOC_NR1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NR1_OUTPUT",DT_DOC_NR1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''''--------TransactionCode-/NF-90----------''''

Call SetTcode(DT_AS01_0105_OKCD) 
Call TakeScreenShot
Call PressEnter()     
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_AS01_0100_POSTING_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0302_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False)
Call PressEnter()
Call GetStatusBar("MessageType", "STATUS_BAR_OUTPUT")
Call PressEnter()
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call PressEnter()
Call SetTextbox("Text","BSEG-SGTXT","",DT_AS01_0305_TEXT,False)
Call PressEnter()
Call ClickButton("Display Document Overview   \(Shift\+F2\)", False)
Call VerifyTextBoxContent("C", "RF05A-AZSAL", "", DT_AS01_0700_CHECK_TEXT_OF_C, False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call PressEnter()
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot
'
'
''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call PressEnter()
Call PressEnter() 

Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV","",ConvertDate(DT_AS01_1142_CHECK_TEXT_OF_CAPITALIZED_ON),False)
Call VerifyTextBoxContent("First acquisition on","ANLA-ZUGDT","",ConvertDate(DT_AS01_1142_CHECK_TEXT_OF_FIRST_ACQUISITION_ON),False)
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("ANLZ-XSTIL", DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN_OCC1)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1)
Call VerifyTableCellContent(3, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2)
Call VerifyTableCellContent(4, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3)
Call VerifyTableCellContent(5, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_4)
'Call VerifyTableCellContent(1, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(2, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_1))
Call VerifyTableCellContent(3, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_2))
Call VerifyTableCellContent(4, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_3))
'Call VerifyTableCellContent(5, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_4))
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)
Call VerifyGridCellContent("Transactions", 1, "Asset Value Date", 0, ConvertDate(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions", 1, "Amount Posted", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 1, "WAERS", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERS)
Call SelectTab("IDC_TABSTRIP", "Posted values", False)
Call TakeScreenShot


'''''''--------TransactionCode-AS02----------''''

Call SetTcode(DT_AS01_1000_OKCD)  
Call PressEnter()
Call PressEnter()  
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call ClickButton("More Intervals", False)
Call ClickButton("Insert Row", False)
Call SetTextbox("From-date of new interval","ANLZ-ADATU","",ConvertDate(DT_AS01_3010_FROMDATE_OF_NEW_INTERVAL),True)
Call TakeScreenShot
Call ClickButton("Yes   \(Enter\)", True)
Call SetTableData("SAPLAISTTIME","Shutdown",1,"<NA>","<NA>",DT_AS01_3000_TABLECELL_SHUTDOWN_0,False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call GetStatusBar("MessageType", "MESSAGE_TYPE_S_OUTPUT")
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT",DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot

''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD_OCC1)  
Call PressEnter()
Call PressEnter() 
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("ANLZ-XSTIL", DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN_OCC2)
Call ClickButton("More Intervals", False)
Call VerifyTableCellContent(1, "Shutdown", "SAPLAISTTIME", DT_AS01_3000_CHECK_SELECTED_OF_TABLECELL_SHUTDOWN_0)
Call VerifyTableCellContent(1, "Fr", "SAPLAISTTIME", ConvertDate(DT_AS01_3000_CHECK_TEXT_OF_TABLECELL_FR_0))
Call VerifyTableCellContent(1, "Cost Center", "SAPLAISTTIME", DT_AS01_3000_CHECK_TEXT_OF_TABLECELL_COST_CENTER_0)
Call ClickButton("Back   \(F3\)", False)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
'Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 5, "VERAENDERUNG", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_VERAENDERUNG)
'Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 5, "JENDE", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE)
Call SelectTab("IDC_TABSTRIP", "Posted values", False)
Call TakeScreenShot
Call ActivateNodeGuiTree(0, "#1;#2")
'Call VerifyGridCellContent("Depreciation posted/planned", 6, "NAFAZ", 0, DT_CHECK_ORD_DEP)
Call TakeScreenShot


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




