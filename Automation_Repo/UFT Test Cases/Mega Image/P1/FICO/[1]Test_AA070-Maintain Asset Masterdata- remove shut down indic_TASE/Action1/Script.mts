
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA070-Maintain Asset Masterdata- remove shut down indic
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


gstrTestCaseName = "Test_AA070-Maintain Asset Masterdata- remove shut down indic"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA070-Maintain Asset Masterdata- remove shut down indic.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-AS01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call TakeScreenShot
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)

Call GetStatusBar("item1", "DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("The asset "&DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT&" 0 is created")

''''''--------TransactionCode-/nF-90----------''''

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0105_OKCD)     
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS01_0100_REFERENCE,False)
Call PressEnter()  

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("C","RF05A-AZSAL","",trim(DT_AS01_0700_CHECK_TEXT_OF_C),False)
Call PressEnter()

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

''''''--------TransactionCode-/nAS11----------''''

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_COMPANY_CODE,False)
Call SetTextbox("Number of similar subnumbers","RA02S-NASSETS","",DT_AS01_SUBNUMBER,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("ANLZ-XSTIL", DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN)
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("ANLZ-XSTIL", DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN_OCC1)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)
Call GetStatusBar("item1", "DT_AS01_0110_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("The asset "&DT_AS01_0110_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" 1 is created")


''''''--------TransactionCode-/nF-90----------''''

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS01_0110_OKCD)     
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT_OCC1,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_AS01_0100_REFERENCE_OCC1,False)
Call PressEnter()
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT_OCC1,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT_OCC1,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE_OCC1,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0302_ACCOUNT_OCC1,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE_OCC1,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT_OCC1,False)
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("C","RF05A-AZSAL","",trim(DT_AS01_0700_CHECK_TEXT_OF_C_OCC1),False)
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot


'''--------TransactionCode-AS02----------''''

Call SetTcode(DT_AS01_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call ClickButton("More Intervals", False)
Call ClickButton("%#AUTOTEXT002", False)  ''Add interval
Call SetTextbox("From-date of new interval","ANLZ-ADATU","",ConvertDate(DT_AS01_3010_FROMDATE_OF_NEW_INTERVAL),True)
Call TakeScreenShot
Call ClickButton("Yes   \(Enter\)", True)
Call SetTableData("SAPLAISTTIME", "Shutdown", 1, "<NA>", "<NA>", DT_AS01_3000_TABLECELL_SHUTDOWN_0, False)
Call SetTableData("SAPLAISTTIME", "Shutdown", 2, "<NA>", "<NA>", DT_AS01_3000_TABLECELL_SHUTDOWN_1, False)
Call ClickButton("btn\[11\]", False)
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2_OUTPUT")
Call VerifyStatusBar("The asset "&DT_AS01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2_OUTPUT&" 1 is changed")
Call TakeScreenShot



''''''--------TransactionCode- /NAS03----------''''

Call SetTcode(DT_AS01_0100_OKCD_OCC2)  
Call TakeScreenShot
Call PressEnter() 
Call PressEnter()

Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call ClickButton("More Intervals", False)

Call VerifyTableCellContent(1, "Shutdown", "SAPLAISTTIME", DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN_OCC2)
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




