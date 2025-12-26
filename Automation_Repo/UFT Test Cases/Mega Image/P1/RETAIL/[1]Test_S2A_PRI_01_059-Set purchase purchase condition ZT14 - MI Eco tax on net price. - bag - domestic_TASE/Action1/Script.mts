
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_059-Set purchase purchase condition ZT14 - MI Eco tax on net price. - bag - domestic
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)





'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_059-Set purchase purchase condition ZT14 - MI Eco tax on net price. - bag - domestic
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRI_01_059- ZT14 - MI Eco  - bag - domestic"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_01_059-Set purchase purchase condition ZT14 - MI Eco tax.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()   
Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","",DT_MM42_0100_PURCHASING_ORG,False)
Call SetTextboxNoLabel("RMMW1-LIFNR","",DT_MM42_0100_VENDOR,False)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)
Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", "Purchasing", False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("For the Conditions",False)
Call ClickButton("New validity period with reference   \(F8\)",True)
Call SetTableData("SAPMV13ATCTRL_D0201", "CnTy", 3, "", "", DT_MM42_0201_TABLECELL_CNTY_2, False)
Call SetTableData("SAPMV13ATCTRL_D0201", "Amount", 3, "", "", DT_MM42_0201_TABLECELL_AMOUNT_2, False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()

Call FinalStatus()





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




