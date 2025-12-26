

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04-02-Article MD- Fresh article BE creation via Matcopy_p2
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

gstrTestCaseName = "Test_04-02-Article MD- Fresh article BE creation via Matcopy_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------VKP5 -----------------------------------


Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","",FormatBlank(DT_VKP5_0100_CREATED_BY),True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog .*","","Variant name",DT_VKP5_0841_SEARCH_TERM,True)
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)

Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","S_PLTYP-LOW","",DT_VKP5_1000_PRICE_LIST,False)
Call SetTextbox("to","S_PLTYP-HIGH","",DT_VKP5_1000_TO,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call TakeScreenShot
Call PressEnter()     ' 
Call PressEnter()     ' 
Call ClickButton("Execute   \(F8\)",False)
Call PressEnter()  
Call TakeScreenShot' 
Call ClickButton("Select All   \(F7\)",False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")


'//-----------------------------------VK11 -----------------------------------

Call SetTcode(DT_VKP5_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VKP5_1000_OKCD)


Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VKP5_0100_CONDITION_TYPE,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRadioButton("RV130-SELKZ", "Sales Org./Dist. Channel/Price List/Article/Sales Unit", True)
Call TakeScreenShot
Call ClickButton("Choose   \(Enter\)",True)
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VKP5_1155_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VKP5_1155_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","KOMG-PLTYP","",DT_VKP5_1155_PRICE_LIST,False)
Call PressEnter()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "", "", DT_VKP5_1155_TABLECELL_ARTICLE_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 1, "", "", DT_VKP5_1155_TABLECELL_SALES_UNIT_0, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "", "", DT_VKP5_1155_TABLECELL_AMOUNT_0, False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(DT_VKP5_1155_CHECK_TEXT_OF_STATUSBAR)

Call PressEnter()
Call SetTextbox("Price List","KOMG-PLTYP","",DT_VKP5_1155_PRICE_LIST_OCC1,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "", "", DT_VKP5_1155_TABLECELL_ARTICLE_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Sales unit", 1, "", "", DT_VKP5_1155_TABLECELL_SALES_UNIT_0_OCC1, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "", "",DT_VKP5_1155_TABLECELL_AMOUNT_0_OCC1, False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(DT_VKP5_1155_CHECK_TEXT_OF_STATUSBAR_OCC1)

 
Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


