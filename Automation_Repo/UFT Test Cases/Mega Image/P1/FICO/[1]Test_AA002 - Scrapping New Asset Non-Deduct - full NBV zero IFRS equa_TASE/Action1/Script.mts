
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA002 - Scrapping New Asset Non-Deduct - full NBV zero IFRS equa
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


gstrTestCaseName = "Test_AA002 - Scrapping New Asset Non-Deduct - full NBV zero IFRS equa"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA002 - Scrapping New Asset Non-Deduct - full NBV zero IFRS equa.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company code","BUKRS-LOW","",DT_S_ALR_87011966_1000_COMPANY_CODE,False)
Call ClickButton("All Selections   \(Shift\+F7\)",False)
Call ClickButton("%_BUKRS_%_APP_%-VALU_PUSH",False)
Call SelectTab("TAB_STRIP", "Exclude Single Values", True)
Call TakeScreenShot
Call ClickCellTable("SAPLALDBSINGLE_E", "Selection Options", 1, "<NA>", "<NA>", True)
Call SelectRowGuiGrid("Company code", "", "Description", "Single Value", True)
'Call ClickButton("Continue   \(Enter\)",True) ' R1E existing code before upgrade
Call ClickButton("Copy   \(Enter\)",True) 'Updated during RGB execution due to property change

Call SetTableData("SAPLALDBSINGLE_E", "Single value", 1, "", "", DT_S_ALR_87011966_3030_TABLECELL_SINGLE_VALUE_0, True)
Call ClickButton("Copy   \(F8\)",True)
Call SelectRadioButton("XEINZEL", "List assets", False)

Call SetTextbox("Sort Variant","SRTVR","",DT_S_ALR_87011966_1000_SORT_VARIANT,False)
Call SetTextbox("Display variant","P_VARI","",DT_S_ALR_87011966_1000_DISPLAY_VARIANT,False)
Call SetTextbox("Capitalization date","SO_AKTIV-LOW","",ConvertDate(DT_S_ALR_87011966_1000_CAPITALIZATION_DATE),False)
Call SetTextbox("to","SO_AKTIV-HIGH","",ConvertDate(DT_S_ALR_87011966_1000_TO),False)
Call TakeScreenShot
Call FocusTextBox("Acquisition value", "SO_KANSW-LOW", False)
Call ClickButton("%_SO_KANSW_%_APP_%-VALU_PUSH",False)
Call SelectTab("TAB_STRIP", "Exclude Single Values", True)
Call TakeScreenShot
Call ClickCellTable("SAPLALDBSINGLE_E", "Selection Options", 1, "<NA>", "<NA>", True)
Call SelectRowGuiGrid("Acquisition value", "", "Description", "Single Value", True)
'Call ClickButton("Continue   \(Enter\)",True) ' R1E existing code before upgrade
Call ClickButton("Copy   \(Enter\)",True)'Updated during RGB execution due to property change
Call ClickButton("Copy   \(F8\)",True)
Call FocusTextBox("Book value", "SO_BCHWR-LOW", False)
Call ClickButton("%_SO_BCHWR_%_APP_%-VALU_PUSH",False)
Call SelectTab("TAB_STRIP", "Exclude Single Values", True)
Call TakeScreenShot
Call ClickCellTable("SAPLALDBSINGLE_E", "Selection Options", 1, "<NA>", "<NA>", True)
Call SelectRowGuiGrid("Book value", "", "Description", "Single Value", True)
'Call ClickButton("Continue   \(Enter\)",True)' R1E existing code before upgrade
Call ClickButton("Copy   \(Enter\)",True)'Updated during RGB execution due to property change
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call GetGridContent("",0,"Asset",3,"<NA>","<NA>","DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN0_OUTPUT")
Call GetGridContent("",0,"Sub-number",3,"<NA>","<NA>","DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN2_OUTPUT")

'''''--------TransactionCode-/nABAVN---------''''

Call SetTcode(DT_S_ALR_87011966_0500_OKCD)     
Call PressEnter()     
Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN0_OUTPUT,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN2_OUTPUT,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_S_ALR_87011966_0200_DOCUMENT_DATE),False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_S_ALR_87011966_0202_ASSET_VALUE_DATE),False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Additional Details", False)
Call TakeScreenShot
Call SetTextbox("Document type","RAIFP1-BLART","",DT_S_ALR_87011966_0204_DOCUMENT_TYPE,False)
Call TakeScreenShot
''Call SetTextbox("Posting period","RAIFP2-MONAT","",DT_S_ALR_87011966_0203_POSTING_PERIOD,False)
Call TakeScreenShot
Call SetTextbox("Transaction Type","RAIFP1-BWASL","",DT_S_ALR_87011966_0205_TRANSACTION_TYPE,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Partial retirement", False)
Call TakeScreenShot

Call ClickButton("Simulate   \(F9\)",False)
Call GetGridContent("Line items",0,"Posting Key",1,"<NA>","<NA>","DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OUTPUT")
Call GetGridContent("Line items",0,"Posting Key",2,"<NA>","<NA>","DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OUTPUT")
Call GetGridContent("Line items",0,"WRBTR",4,"<NA>","<NA>","DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR_OUTPUT")
Call VerifyGridCellContent("Line items", 4, "WRBTR", 0, DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR_OUTPUT)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

''''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_S_ALR_87011966_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call SetTextbox("Asset","ANLA-ANLN1","",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN0_OUTPUT,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANLN2_OUTPUT,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_S_ALR_87011966_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 11, "Value", 0, lcase(DT_S_ALR_87011966_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_BEZEICHNUNG))
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 11, "Year-end", 0, "")
Call ClickButtonIfExist("Exit \(Shift\+F3\)",False)
Call ClickButtonIfExist("Exit \(Shift\+F3\)",False)


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




