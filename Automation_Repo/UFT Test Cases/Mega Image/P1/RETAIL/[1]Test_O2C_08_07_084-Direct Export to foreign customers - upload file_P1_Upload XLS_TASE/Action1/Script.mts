'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_084-Direct Export to foreign customers - upload file_P1_Upload XLS     
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



gstrTestCaseName = "Test_O2C_08_07_084-Direct Export to foreign customers - upload file_P1_Upload XLS"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_084-Direct Export to foreign customers - upload file_P1_Upload XLS.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-ZMDIM_STO_SO_FILE----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDIM_STO_SO_FILE_0841_SEARCH_TERM,True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call SelectRowGuiGrid("Variant Catalog for Program ZMDIM_CREATE_PO_STO_FROM_FILE"&".*", 0, "Variant name", DT_ZMDIM_STO_SO_FILE_0841_SEARCH_TERM, True)
'Call SelectRowGuiGrid("Variant Catalog for Program ZMDIM_CREATE_PO_STO_FROM_FILE", 0, "Variant name", DT_ZMDIM_STO_SO_FILE_0841_SEARCH_TERM, True)
Call ClickButton("Choose   \(F2\)",True)
Call SetTextbox("Sales Organizat","P_VKORG","",DT_ZMDIM_STO_SO_FILE_1000_SALES_ORGANIZAT,False)
Call SetTextbox("Distribution Channel","P_VTWEG","",DT_ZMDIM_STO_SO_FILE_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","P_SPART","",DT_ZMDIM_STO_SO_FILE_1000_DIVISION,False)
Call SetTextbox("Sales document type","P_AUART","",DT_ZMDIM_STO_SO_FILE_1000_SALES_DOCUMENT_TYPE,False)
Call SetTextbox("Sold-to party","P_KUNNR","",DT_ZMDIM_STO_SO_FILE_1000_SOLDTO_PARTY,False)
Call SetTextbox("Ship to Party","P_KUNWE","",DT_ZMDIM_STO_SO_FILE_1000_SHIP_TO_PARTY,False)
Call SetTextbox("Site","P_SOWRK","",DT_ZMDIM_STO_SO_FILE_1000_SITE,False)
Call SetTextbox("Storage location","P_SOLGO","",DT_ZMDIM_STO_SO_FILE_1000_STORAGE_LOCATION,False)
Call SetComboByKey("Type of Article Code", DT_ZMDIM_STO_SO_FILE_1000_TYPE_OF_ARTICLE_CODE)
Call SelectRadioButton("P_CRSO", "Sales Inquiry/Order Creation", False)
Call SelectRadioButton("P_MEINS", "Use Base UoM", False)
Call TakeScreenShot
Call SetTextbox("File name","P_FILE","","",False)
Call FocusTextBox("File name", "P_FILE", False)
Call SendKey("{F4}")
Wait(5)
Call SetTextbox("Directory","DY_PATH","",DT_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_FILENAME,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'''Call GetGridContent("", 0, "Message Text", 4, "<NA>", "<NA>", "DT_ZMDIM_STO_SO_FILE_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3_T_MSG_OUTPUT")
Call GetGridContent("", 0, "Message Text", 5, "<NA>", "<NA>", "DT_ZMDIM_STO_SO_FILE_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3_T_MSG_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''Call VerifyGridCellContent("", 4, "Message Text", 0, lcase(DT_ZMDIM_STO_SO_FILE_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3_T_MSG_OCC1))
Call VerifyGridCellContent("", 5, "Message Text", 0, lcase(DT_ZMDIM_STO_SO_FILE_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_3_T_MSG_OCC1))
Call TakeScreenShot
''''--------TransactionCode-/nVA02----------''''

Call SetTcode(DT_TRANSACTION2)     
Call PressEnter()   
Call TakeScreenShot
Call SetTextbox("Order","VBAK-VBELN","",DT_SALES_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter() 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "First date", 1, "", "", ConvertDate(DT_DATE), False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call PressEnter() 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG", "First date", 2, "", "", ConvertDate(DT_DATE), False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Save   \(Ctrl\+S\)",False)
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



