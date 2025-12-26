
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA051 Reverse Capitalization
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


gstrTestCaseName = "Test_AA051 Reverse Capitalization"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA051 Reverse Capitalization.xls"
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

Call SetTextbox("Company Code","RLAB01-BUKRS","",DT_AB08_0010_COMPANY_CODE,False)
Call SetTextbox("Asset","RLAB01-ANLN1","",DT_AB08_0010_ASSET,False)
Call SetTextboxNoLabel("RLAB01-ANLN2","",DT_AB08_0010_SUBNUMBER,False)
Call SetTextbox("Fiscal Year","RLAB01-GJAHR","",DT_AB08_0010_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
''Call SelectRowGuiTable("SAPLAB01TC_ZAB01_D0100", columnName, cellVal, blnIsItPopup)
Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100", 1, False)
Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100", 2, False)
Call ClickButton("Reverse Document   \(F6\)", False)
Call TakeScreenShot
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_AB08_0100_DOCUMENT_TYPE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AB08_0010_ASSET,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Line items", 5, "WRBTR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_WRBTR)
Call VerifyGridCellContent("Line items", 1, "WRBTR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call ClickButton("Back   \(F3\)", False)
Call PressEnter()     
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100", 1, False)
Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100", 2, False)
Call ClickButton("Reverse Document   \(F6\)", False)
Call TakeScreenShot

Call SelectCheckbox("LKO74-TESTLAUF", 0, "OFF", False)
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_AB08_0100_DOCUMENT_TYPE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Line items", 5, "WRBTR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_WRBTR_OCC1)
Call VerifyGridCellContent("Line items", 1, "WRBTR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR_OCC1)
Call VerifyGridCellContent("Line items", 5, "WAERS", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_WAERS)
Call VerifyGridCellContent("Line items", 1, "GSBER", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call GetStatusBar("Text", "DT_AB08_0010_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar(DT_AB08_0010_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT)


''''''--------TransactionCode- /NAS03----------''''

Call SetTcode(DT_AB08_0010_OKCD)  
Call PressEnter() 
Call TakeScreenShot
Call PressEnter() 

Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Transactions", 3, "BWASL", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BWASL)
Call VerifyGridCellContent("Transactions", 3, "BWATXT", 0, lcase(DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BWATXT))
Call VerifyGridCellContent("Transactions", 3, "WAERS", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WAERS)
Call GetGridContent("Planned values IFRS APC, depreciation", "", "JENDE", 4, "<NA>", "<NA>", "DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OUTPUT")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", "", DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OUTPUT)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AB08_0100_ASSET,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", "", "")
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL)
Call VerifyGridCellContent("Transactions", 2, "BWATXT", 0, lcase(DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWATXT))
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AB08_0100_ASSET_OCC1,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", "", "")
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC1)
Call VerifyGridCellContent("Transactions", 2, "BWATXT", 0, lcase(DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWATXT_OCC1))
Call VerifyGridCellContent("Transactions", 2, "WAERS", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WAERS)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AB08_0100_ASSET_OCC2,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", "", "")
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC2)
Call VerifyGridCellContent("Transactions", 2, "BWATXT", 0, lcase(DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWATXT_OCC2))
Call VerifyGridCellContent("Transactions", 2, "WAERS", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WAERS_OCC1)



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




