
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_03_15-Create sales order to domestic customer-VK11 
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


gstrTestCaseName = "Test_S2A_PRI_03_15-Create sales order to domestic customer-VK11"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRI_03_15-Create sales order to domestic customer-VK11.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''''-----Login----------'''''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-VK11----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1970_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1970_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Customer","KOMG-KUNNR","",DT_VK11_1970_CUSTOMER,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_AMOUNT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Unit", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_UNIT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "per", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_PER_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "UoM", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_UOM_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Calculat.type", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_CALCULATTYPE_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Valid From", 1, "<NA>", "<NA>", ConvertDate(DT_VK11_1970_TABLECELL_VALID_FROM_0), False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Valid to", 1, "<NA>", "<NA>", ConvertDate(DT_VK11_1970_TABLECELL_VALID_TO_0), False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("btn\[0\]",True)
Call ClickButtonIfExist("btn\[0\]",True)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot
Call VerifyStatusBar(lcase(DT_VERIFY_SAVED_MESSAGE))
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_ARTICLE_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_AMOUNT_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Unit", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_UNIT_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "per", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_PER_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "UoM", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_UOM_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Calculat.type", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_CALCULATTYPE_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Valid From", 1, "<NA>", "<NA>", ConvertDate(DT_VK11_1970_TABLECELL_VALID_FROM_0_OCC1), False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Valid to", 1, "<NA>", "<NA>", ConvertDate(DT_VK11_1970_TABLECELL_VALID_TO_0_OCC1), False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(lcase(DT_VERIFY_SAVED_MESSAGE))
Call TakeScreenShot

'''''''--------TransactionCode-/nVK11----------''''

Call SetTcode(DT_VK11_1970_OKCD)     
Call PressEnter()     

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE_OCC1,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VK11_1970_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1970_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Customer","KOMG-KUNNR","",DT_VK11_1970_CUSTOMER_OCC1,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_ARTICLE_0_OCC2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "<NA>", "<NA>", DT_VK11_1970_TABLECELL_AMOUNT_0_OCC2, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Valid From", 1, "<NA>", "<NA>", ConvertDate(DT_VK11_1970_TABLECELL_VALID_FROM_0_OCC2), False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Valid to", 1, "<NA>", "<NA>", ConvertDate(DT_VK11_1970_TABLECELL_VALID_TO_0_OCC2), False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(lcase(DT_VERIFY_SAVED_MESSAGE))
Call TakeScreenShot

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




