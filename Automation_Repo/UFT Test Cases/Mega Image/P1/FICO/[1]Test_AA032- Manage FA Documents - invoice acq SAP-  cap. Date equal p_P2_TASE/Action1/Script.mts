
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA032- Manage FA Documents - invoice acq SAP-  cap. Date equal p_P2
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


gstrTestCaseName = "Test_AA032- Manage FA Documents - invoice acq SAP-  cap. Date equal p_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA032- Manage FA Documents - invoice acq SAP-  cap. Date equal p_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
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

Call SetTextbox("Document Number","RF05L-BELNR","",DT_AS11_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_AS11_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_AS11_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

Call VerifyGridCellContent("", 1, "RF05A_UBAZW", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "RF05A_UBAZW", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 3, "RF05A_UBAZW", 0, DT_AS11_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW)

'''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS11_0750_OKCD)  
Call PressEnter()     
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS11_0100_ASSET,False)
Call TakeScreenShot
Call ClickButton("Master data   \(F7\)", False)
Call SelectTab("TABSTRIP100", "General", False)
Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV","",ConvertDate(DT_AS11_1142_CHECK_TEXT_OF_CAPITALIZED_ON),False)
Call VerifyTextBoxContent("First acquisition on","ANLA-ZUGDT","",ConvertDate(DT_AS11_1142_CHECK_TEXT_OF_FIRST_ACQUISITION_ON),False)
Call SelectTab("TABSTRIP100", "Origin", False)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0_OCC1)
Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1_OCC1)
Call VerifyTableCellContent(3, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2_OCC1)
Call VerifyTableCellContent(4, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3_OCC1)
Call VerifyTableCellContent(5, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_4_OCC1)
Call VerifyTableCellContent(1, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(2, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(3, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(4, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(5, "ODep Start", "SAPLAISTTC_ANLB", ConvertDate(DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call VerifyGridCellContent("Transactions", 1, "Asset Value Date", 0, ConvertDate(DT_AS11_1142_CHECK_TEXT_OF_CAPITALIZED_ON))
Call VerifyGridCellContent("Transactions", 1, "Amount Posted", 0, DT_AS11_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "Transaction Type", 0, DT_AS11_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)


Call SelectTab("IDC_TABSTRIP", "Parameters", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR","",DT_AS11_0304_CHECK_TEXT_OF_USEFUL_LIFE,False)
Call VerifyTextBoxContent("Remaining life","AW01_DEP_PAR-RSTNDJ","",DT_AS11_0304_CHECK_TEXT_OF_REMAINING_LIFE,False)
Call VerifyTextBoxContent("Ord.dep.start date","AW01_DEP_PAR-AFABG","",ConvertDate(DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0),False)
Call ClickButton("Display Master Data   \(Shift\+F4\)", False)
Call SelectTab("TABSTRIP100", "Time-dependent", False)
'Call VerifyTextBoxContent("Asset shutdown","ANLZ-XSTIL","","",False)
Call VerifyCheckBoxValue("ANLZ-XSTIL","off" )

Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Exit   \(Shift\+F3\)", False)
Call ClickButton("Exit   \(Shift\+F3\)", False)


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




