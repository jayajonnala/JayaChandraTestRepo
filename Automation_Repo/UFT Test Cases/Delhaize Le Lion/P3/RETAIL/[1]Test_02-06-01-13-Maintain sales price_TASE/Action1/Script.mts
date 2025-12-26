
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02-06-01-13-Maintain sales price_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "TC1_Test_02-06-01-13-Maint sale"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_PRE_Get_Art_Doc_with_Reference_MB51.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''--------TransactionCode-VK11----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextBox("Condition Type","RV13A-KSCHL","",DT_VK11_0100_CONDITION_TYPE,False)
Call PressEnter()
Call TakeScreenShot

Call SelectRadioButton("RV130-SELKZ","Sales Org\./Dist\. Channel/Price List/Article/Sales Unit",True)
Call TakeScreenShot
Call ClickButton("Choose   \(Enter\)",True)
Call TakeScreenShot

Call SetTextBox("Sales Organization","KOMG-VKORG","",DT_VK11_1155_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","KOMG-VTWEG","",DT_VK11_1155_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("Price List","KOMG-PLTYP","",DT_VK11_1155_PRICE_LIST,False)
Call TakeScreenShot
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Article","1","","",DT_VK11_1155_TABLECELL_ARTICLE_0,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Sales Unit","1","","",DT_VK11_1155_TABLECELL_SALES_UNIT_0,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount","1","","",DT_VK11_1155_TABLECELL_AMOUNT_0,False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_VK11_1155_CHECK_TEXT_OF_STATUSBAR)
Call ClickButtonIfExist("Enter   \(F5\)",True)

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


