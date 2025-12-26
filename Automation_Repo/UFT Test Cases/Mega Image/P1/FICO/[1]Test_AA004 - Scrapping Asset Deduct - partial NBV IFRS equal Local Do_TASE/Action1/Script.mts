
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA004 - Scrapping Asset Deduct - partial NBV IFRS equal Local Do
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



gstrTestCaseName = "Test_AA004 - Scrapping Asset Deduct - partial NBV IFRS equal Local Do"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA004 - Scrapping Asset Deduct - partial NBV IFRS equal Local Do.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_OF_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
If CDate(DT_LAST_DATE_OF_PERIOD_OCC1)<=CDate(DT_LAST_DATE_OF_PERIOD) Then
	Call WriteRunTimeDataToExcelGlobalSheet ("DT_DATE",DT_LAST_DATE_OF_PERIOD_OCC1)
Else
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DATE",DT_LAST_DATE_OF_PERIOD)
End If
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''--------------login----------------'''''
'
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Company code","BUKRS-LOW","",DT_S_ALR_87011966_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call ClickButton("All Selections   \(Shift\+F7\)",False)
Call SelectRadioButton("XEINZEL", "List assets", False)
Call SetTextbox("Asset class","SO_ANLKL-LOW","",DT_S_ALR_87011966_1000_ASSET_CLASS,False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_S_ALR_87011966_1000_DEPRECIATION_AREA,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_S_ALR_87011966_1000_SORT_VARIANT,False)
Call SetTextbox("Capitalization date","SO_AKTIV-LOW","",ConvertDate(DT_S_ALR_87011966_1000_CAPITALIZATION_DATE),False)
Call SetTextbox("to","SO_AKTIV-HIGH","",ConvertDate(DT_S_ALR_87011966_1000_TO),False)
Call TakeScreenShot
Call FocusTextBox("Book value", "SO_BCHWR-LOW", False)
Call ClickButton("%_SO_BCHWR_%_APP_%-VALU_PUSH",False)
Call ClickCellTable("SAPLALDBSINGLE", "Selection Options", 1, "<NA>", "<NA>", True)
Call SelectRowGuiGrid("Book value", "", "Description", "Greater than", True)
Call ClickButton("btn\[0\]",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_S_ALR_87011966_3010_TABLECELL_SINGLE_VALUE_0, True)
Call PressEnter()
Call ClickButton("btn\[0\]",True)
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
'Code
Call GetGridContent("",0,"Asset",4,"<NA>","<NA>","DT_S_ALR_1_ANLN0_OUTPUT")
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANLN0_OUTPUT",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANLN0)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetGridContent("",0,"Sub-number",4,"<NA>","<NA>","DT_S_ALR_1_ANLN2_OUTPUT")
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANLN2_OUTPUT",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANLN2)
''Call GetGridContent("",0,"Asset",2,"<NA>","<NA>","DT_S_ALR_1_ANLN0_OUTPUT_1")
''Call GetGridContent("",0,"Sub-number",2,"<NA>","<NA>","DT_S_ALR_1_ANLN2_OUTPUT_1")
''Call GetGridContent("",0,"Asset",3,"<NA>","<NA>","DT_S_ALR_1_ANLN0_OUTPUT_2")
''Call GetGridContent("",0,"Sub-number",3,"<NA>","<NA>","DT_S_ALR_1_ANLN2_OUTPUT_2")
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''--------TransactionCode-/NS_ALR_87012052----------''''

Call SetTcode(DT_S_ALR_87011966_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company code","BUKRS-LOW","",DT_S_ALR_87011966_1000_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call ClickButton("All Selections   \(Shift\+F7\)",False)
Call SelectRadioButton("XEINZEL", "List assets", False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_S_ALR_87011966_GRIDCELL_1_ANLN0,False)
Call SetTextbox("Subnumber","UNTNR-LOW","",DT_S_ALR_87011966_GRIDCELL_1_ANLN2,False)
Call SetTextbox("Asset class","SO_ANLKL-LOW","",DT_S_ALR_87011966_1000_ASSET_CLASS_OCC1,False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_S_ALR_87011966_1000_DEPRECIATION_AREA_OCC1,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_S_ALR_87011966_1000_SORT_VARIANT_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
'tempVal=0
''While CheckifGuiLabelExists("Asset")=True
'	tempVal=tempVal+1
'	'Call ClickButton("Back   \(F3\)",False)
'	AssetVal="DT_S_ALR_87011966_GRIDCELL_1_ANLN0"&"_"&tempVal
'	Call SetTextbox("Asset number","ANLAGE-LOW","",AssetVal,False)
'	Call SetTextbox("Subnumber","UNTNR-LOW","",DT_S_ALR_87011966_GRIDCELL_1_ANLN2&"_"&tempVal,False)
'	Call ClickButton("Execute   \(F8\)",False)
''Wend 
'Set tempVal=nothing
Call VerifyTextBoxContent("Information Message", "MESSTXT1", "", lcase(DT_S_ALR_87011966_0010_CHECK_TEXT_OF_MESSTXT1_OCC1), True)
Call ClickButton("Continue   \(Enter\)",True)
'''Call WriteRunTimeDataToExcelGlobalSheet ("DT_Count",Cint(DT_Count)+1)

'''--------TransactionCode-/nABAVN---------''''
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_S_ALR_87011966_1000_OKCD)     
Call PressEnter()     
Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_S_ALR_87011966_0300_ASSET,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_S_ALR_87011966_0300_ASSET_OCC1,False)
Call TakeScreenShot
Call PressEnter()
DT_S_ALR_87011966_0200_DOCUMENT_DATE_OCC1 = "15."&ConvertDoubledigit(Cstr(MONTH(Date())))&"."&YEAR(Date())
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_S_ALR_87011966_0200_DOCUMENT_DATE_OCC1),False)
Call TakeScreenShot
Call PressEnter()

''DT_S_ALR_87011966_0202_ASSET_VALUE_DATE_OCC1= "30."&ConvertDoubledigit(Cstr(MONTH(Date())))&"."&YEAR(Date())
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_S_ALR_87011966_0200_DOCUMENT_DATE_OCC1),False)
Call TakeScreenShot
Call PressEnter()

''DT_S_ALR_87011966_0201_POSTING_DATE_OCC1= "30."&ConvertDoubledigit(Cstr(MONTH(Date())))&"."&YEAR(Date())
''Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_S_ALR_87011966_0201_POSTING_DATE),False)
Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_DATE),False)
Call TakeScreenShot
Call PressEnter()



Call SelectTab("TABSTRIP100", "Additional Details", False)
Call TakeScreenShot
Call SetTextbox("Document type","RAIFP1-BLART","",DT_S_ALR_87011966_0204_DOCUMENT_TYPE,False)
Call TakeScreenShot
Call SetTextbox("Posting period","RAIFP2-MONAT","",DT_S_ALR_87011966_0203_POSTING_PERIOD,False)
Call TakeScreenShot
Call SetTextbox("Transaction Type","RAIFP1-BWASL","",DT_S_ALR_87011966_0205_TRANSACTION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Partial retirement", False)
Call TakeScreenShot
Call SetTextbox("Amount posted","RAIFP2-ANBTR","",DT_S_ALR_87011966_0401_AMOUNT_POSTED,False)
Call TakeScreenShot
Call ClickButton("Simulate   \(F9\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item2", "DT_AssetTransactionDocument_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. RO02 "&DT_AssetTransactionDocument_OUTPUT)

''''''--------TransactionCode-FB03----------''''

Call SetTcode(DT_S_ALR_87011966_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_AssetTransactionDocument_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_S_ALR_87011966_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_S_ALR_87011966_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
''''Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_S_ALR_87011966_0200_DOCUMENT_DATE_OCC1), False)
''''Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_S_ALR_87011966_0201_POSTING_DATE_OCC1), False)

Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_S_ALR_87011966_0750_CHECK_TEXT_OF_DOCUMENT_DATE), False)
'''Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_S_ALR_87011966_0750_CHECK_TEXT_OF_POSTING_DATE), False)
Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_DATE), False)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Account", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call GetGridContent("", 0, "Amount", 1, "<NA>", "<NA>", "DT_AMOUNT_TO_VALIDATE_OUTPUT")
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_AMOUNT_TO_VALIDATE_OUTPUT",DT_AMOUNT_TO_VALIDATE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''''Call SetGridData("", 1, "Posting Key", cellValue, blnIsItPopup)


''''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_S_ALR_87011966_0750_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call SetTextbox("Asset","ANLA-ANLN1","",DT_S_ALR_87011966_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_S_ALR_87011966_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_S_ALR_87011966_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)

Call ClickButtonToolbar("&MB_FILTER",1)
Call SelectCellGuiGrid("Column Set",0,2,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Asset Value Date","%%DYN001-LOW","",DT_S_ALR_87011966_0202_ASSET_VALUE_DATE,True)
Call ClickButton("Execute   \(Enter\)",True)

''Call VerifyGridCellContent("Transactions", 1, "Asset Value Date", 0, ConvertDate(DT_S_ALR_87011966_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions", 1, "Asset Value Date", 0, DT_S_ALR_87011966_0200_DOCUMENT_DATE_OCC1)
Call VerifyGridCellContent("Transactions", 1, "Amount posted", 0, DT_S_ALR_87011966_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_S_ALR_87011966_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call TakeScreenShot
Call ClickButtonIfExist("Exit \(Shift\+F3\)",False)
Call ClickButtonIfExist("Exit \(Shift\+F3\)",False)

Call LogOff()

Call FinalStatus ()

' ConvertDoubledigit(StrSingle)








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



