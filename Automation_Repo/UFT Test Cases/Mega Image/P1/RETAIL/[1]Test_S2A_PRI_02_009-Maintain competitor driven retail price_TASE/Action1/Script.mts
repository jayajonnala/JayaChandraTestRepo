
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_009-Maintain competitor driven retail price
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
'.................Test Script Name :Test_S2A_PRI_02_009-Maintain competitor driven retail price
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_S2A_PRI_02_009-Maintain  retail price_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_009-Maintain competitor driven retail price.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()   

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_ZMDPC_UPLOAD_COND_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call SelectRowGuiGrid("Variant Catalog for Program ZMDPC_UPLOAD_PRICES_FROM_FILE"&".*", 0, "Variant name", DT_ZMDPC_UPLOAD_COND_0841_SEARCH_TERM, True)
Call ClickButton("Choose   \(F2\)",True)
Call SetTextbox("File name","P_FILE","","",False)
Call FocusTextBox("File name", "P_FILE", False)
Call SendKey("{F4}")
Wait(5)
Call SetTextbox("Directory","DY_PATH","",DT_ZMDPC_UPLOAD_COND_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_ZMDPC_UPLOAD_COND_0200_FILE_NAME,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Select All   \(F5\)",False)
Call ClickButton("Create Conditions   \(F8\)",False)
Call TakeScreenShot

' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Exception",0,DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT)
'Call GetGridContent("", 0, "Exception", 1, "<NA>", "<NA>", "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT_OUTPUT")
Call GetGridContent("", 0, "Article", 1, "<NA>", "<NA>", "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")
Call GetGridContent("", 0, "Customer", 1, "<NA>", "<NA>", "DT_ZMDPC_UPLOAD_COND_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KUNNR_OUTPUT")
'
'
''''--------TransactionCode-/nVK13----------''''

Call SetTcode(DT_ZMDPC_UPLOAD_COND_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_ZMDPC_UPLOAD_COND_0100_CONDITION_TYPE,False)
Call ClickButton("Condition Information   \(Shift\+F4\)",False)
Call SetTextbox("from / on","SEL_DATE","",ConvertDate(DT_ZMDPC_UPLOAD_COND_1000_FROM__ON),False)
Call SetTextbox("Distribution Channel","F002-LOW","",DT_ZMDPC_UPLOAD_COND_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Sales Organization","F001-LOW","",DT_ZMDPC_UPLOAD_COND_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Customer","F006-LOW","",DT_ZMDPC_UPLOAD_COND_1000_CUSTOMER,False)
Call SetTextbox("Article","F004-LOW","",DT_ZMDPC_UPLOAD_COND_1000_ARTICLE,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Select All   \(F7\)",False)
Call ClickButton("Display   \(F5\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Article", "SAPMV13ATCTRL_FAST_ENTRY", DT_ZMDPC_UPLOAD_COND_1153_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1, "Customer", "SAPMV13ATCTRL_FAST_ENTRY", DT_ZMDPC_UPLOAD_COND_1153_CHECK_TEXT_OF_TABLECELL_CUSTOMER_0)
Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_FAST_ENTRY", DT_ZMDPC_UPLOAD_COND_1153_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)

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




