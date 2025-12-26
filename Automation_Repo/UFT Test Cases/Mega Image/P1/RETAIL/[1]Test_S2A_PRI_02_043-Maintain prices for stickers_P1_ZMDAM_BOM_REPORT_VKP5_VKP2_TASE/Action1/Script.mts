
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_02_043-Maintain prices for stickers_P1_ZMDAM_BOM_REPORT_VKP5_VKP2 
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

gstrTestCaseName = "Test_S2A_PRI_02_043-P1_VKP5_VKP2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_043-Maintain prices for stickers_P1_ZMDAM_BOM_REPORT_VKP5_VKP2_ZACTF_LEG_CODES.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
'
'''''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''--------TransactionCode-ZMDAM_BOM_REPORT----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Purchasing group","S_WEKGR-LOW","",DT_ZMDAM_BOM_REPORT_1000_PURCHASING_GROUP,False)
Call SetTextbox("Link Type","S_STLAN-LOW","",DT_ZMDAM_BOM_REPORT_1000_LINK_TYPE,False)
Call SetTextbox("Retention Level","S_STLAL-LOW","",DT_ZMDAM_BOM_REPORT_1000_RETENTION_LEVEL,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "Header Article", false)
Call ClickButton("Set filter   \(Ctrl\+F5\)",false)
Call SetTextbox("Header Article","%%DYN001-LOW","",DT_ARTICLE_INP,True)
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call SelectRowGuiGrid("", 0, "Retention Level", "RO", False)
Call DoubleClickGuiGridCell("", 0, 1, "Item Article", False)
Call TakeScreenShot
Call GetTextboxValue("P_MATNR", 0, "DT_ZMDAM_BOM_REPORT_0100_GET_ARTICLE_OUTPUT", False)
Call GetGridContent("Add- Change BOM", 0, "Parent Art", 1, "<NA>", "<NA>", "DT_ZMDAM_BOM_REPORT_0100_GET_PARENT_ARTICLE_OUTPUT")
Call GetGridContent("Add- Change BOM", 0, "Quantity", 1, "<NA>", "<NA>", "DT_ZMDAM_BOM_REPORT_0100_GET_QUANTITY_OUTPUT")
'
'''''''''--------TransactionCode-VKP5----------''''
'
Call SetTcode(DT_ZMDAM_BOM_REPORT_0100_OKCD) 
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("btn\[17\]",False)
Call SetTextbox("Variant","V-LOW","",DT_ZMDAM_BOM_REPORT_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAM_BOM_REPORT_0100_GET_PARENT_ARTICLE_OUTPUT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot
Call SetGridData("", 1, "ENDPR", DT_ZMDAM_BOM_REPORT_0100_GRIDCELL_0_FINAL_PRICE, False)
Call PressEnter()     
Call TakeScreenShot
Call GetGridContent("", 0, "ENDPR", 1, "<NA>", "<NA>", "DT_ZMDAM_BOM_REPORT_0100_GET_GRIDCELL_0_FINAL_PRICE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call GetStatusBar("MessageType", "DT_VKP5_1000_GET_MESSAGE_TYPE_OUTPUT")
Call GetStatusBar("item1", "DT_ZMDAM_BOM_REPORT_1000_GET_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "&DT_ZMDAM_BOM_REPORT_1000_GET_PRICING_DOCUMENT_OUTPUT&" created")
Call TakeScreenShot

''''''''--------TransactionCode-VK13 below steps added based on workshop on the defect 20267----------''''
Call SetTcode(DT_ZMDAM_BOM_REPORT_0200_OKCD) 
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_CONDITION_TYPE,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Article","F001-LOW","",DT_ARTICLE_INP,False)
Call TakeScreenShot
Call ClickButton("Condition Info   \(Shift\+F4\)",False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0,"ON",False)
Call ClickButton("Display   \(F5\)",False)
Call TakeScreenShot
Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "Article", DT_ARTICLE_INP, "DT_ZSTI_CONDITION_RATE_OUTPUT", False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

last=Cstr( Right(DT_FINAL_PRICE_1,1))
If last="-" Then
	DT_ZSTI=left(DT_FINAL_PRICE_1,len(DT_FINAL_PRICE_1)-1)
	DT_FINAL_PRICE_OUTPUT=DT_CONDITION_RATE-((DT_ZSTI/100)*DT_CONDITION_RATE)
Else
	DT_FINAL_PRICE_OUTPUT=DT_CONDITION_RATE+((DT_ZSTI/100)*DT_CONDITION_RATE)
End If

''''''''--------TransactionCode-VKP2----------''''

Call SetTcode(DT_ZMDAM_BOM_REPORT_1000_OKCD) 
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAM_BOM_REPORT_0100_GET_ARTICLE_OUTPUT,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_ZMDAM_BOM_REPORT_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_ZMDAM_BOM_REPORT_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Validity Period","S_DATUM-LOW","",ConvertDate(DT_ZMDAM_BOM_REPORT_1000_VALIDITY_PERIOD),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
''Call VerifyGridCellContent("Sales Price Conditions", 1, "KBETR", 0, DT_ZMDAM_BOM_REPORT_0100_GET_GRIDCELL_0_FINAL_PRICE_OUTPUT)
Call VerifyGridCellContent("Sales Price Conditions", 1, "KBETR", 0, DT_FINAL_PRICE_OUTPUT)

'''''''''--------TransactionCode-/nZACTF_LEG_CODES----------''''

Call SetTcode(DT_ZMDAM_BOM_REPORT_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Group Name","P_GROUP","",DT_ZMDAM_BOM_REPORT_1000_GROUP_NAME,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAM_BOM_REPORT_0100_GET_ARTICLE_OUTPUT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call GetGridContent("", 0, "Legacy Value", 1, "<NA>", "<NA>", "DT_ZMDAM_BOM_REPORT_0500_GET_GRIDCELL_0_LEGACY_VALUE_OUTPUT")
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


