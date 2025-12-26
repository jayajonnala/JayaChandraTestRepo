
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_207-BOM - stiker
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

gstrTestCaseName = "Test_MD_01_01_207-BOM - stiker_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_MD_01_01_207-BOM - stiker.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
''
'''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''--------TransactionCode-ZMDAM_BOM_DATA_ENTRY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","P_MATNR","",DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE,false)
Call SetTextbox("Link Type","P_STLAN","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LINK_TYPE,false)
Call SetTextbox("Link retention level","P_STLAL","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LINK_RETENTION_LEVEL,false)
Call SetTextbox("Layout","P_VARI","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LAYOUT,false)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)

Call ClickButtonToolBar("CHANGE", 0)
SAPGuiSession("Session").SAPGuiWindow("DG: BOM Update").InsightObject("InsightObject").Click
'
Call SetGridData("Change/Add Item to BOM", 1, "Parent Art", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_PARENT_ART, False)
Call SetGridData("Change/Add Item to BOM", 1, "Quantity", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_QUANTITY, False)
Call SetGridData("Change/Add Item to BOM", 1, "Main Paren", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_MAIN_PAREN, False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButtonToolBar("SAVE", 0)
Call ClickButtonIfExist("Yes", True)
Call TakeScreenShot
Call GetLabelContentByRefLabel("Message text", 0, -32, "DT_ZMDAM_BOM_DATA_ENTRY_0120_CHECK_TEXT_BOM_CHANGED_OUTPUT", True)
Call VerifyifGuiLabelExists(DT_ZMDAM_BOM_DATA_ENTRY_0120_CHECK_TEXT_BOM_CHANGED_OUTPUT)
Call ClickButtonIfExist("Copy   \(Enter\)", True)
Call TakeScreenShot
Call ClickButtonifExist("Execute   \(F8\)", False)

'''Call VerifyGridCellContent("Add- Change BOM", 1, "Parent Art", "", DT_ZMDAM_BOM_DATA_ENTRY_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
''''Call VerifyGridCellContent("Add- Change BOM", 1, "Parent Art", "", DT_ZMDAM_BOM_CHECK_GRIDCELL_0_MATNR)
'
''''--------TransactionCode-MM43----------''''

Call SetTcode(DT_ZMDAM_BOM_DATA_ENTRY_0001_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Article","RMMW1-MATNR","",DT_ZMDAM_BOM_DATA_ENTRY_0100_ARTICLE,false)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","","",false)
'Call SetTextbox("Vendor","RMMW1-LIFNR","","",false)
Call SetTextboxNoLabel("RMMW1-LIFNR","","",False)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_ZMDAM_BOM_DATA_ENTRY_0100_SALES_ORG,false)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_ZMDAM_BOM_DATA_ENTRY_0100_DISTR_CHANNEL,false)

Call ClickButton("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100", "Screen description", DT_ZMDAM_BOM_DATA_ENTRY_4008_SALES, false)
Call TakeScreenShot
Call PressEnter()
'''Call PressEnter()

Call TakeScreenShot
Call ClickButtonIfExist("No",True)
Call TakeScreenShot
'''Call VerifyTextBoxContent("Pr. fixing", "MVKE-PBIND", "", DT_ZMDAM_BOM_DATA_ENTRY_2231_CHECK_TEXT_OF_PR_FIXING, False)

''--------TransactionCode-VK11----------''''

Call SetTcode(DT_ZMDAM_BOM_DATA_ENTRY_4030_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_ZMDAM_BOM_DATA_ENTRY_0100_CONDITION_TYPE,false)
Call TakeScreenShot
'''Call ClickButton("Key Combination   \(Shift\+F5\)",False)
Call PressEnter()
Call TakeScreenShot

Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Calc. Type", 1, "", "", DT_ZMDAM_BOM_DATA_ENTRY_1960_TABLECELL_CALC_TYPE_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "", "", DT_ZMDAM_BOM_DATA_ENTRY_1960_TABLECELL_AMOUNT_0,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call VerifyStatusBar(DT_ZMDAM_BOM_DATA_ENTRY_1960_CHECK_TEXT_OF_STATUSBAR)


'''--------TransactionCode-VKP5----------''''

Call SetTcode(DT_ZMDAM_BOM_DATA_ENTRY_1960_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_ZMDAM_BOM_DATA_ENTRY_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot

Call SetTextbox("Sales price determination seq\.","P_VKERV","",DT_ZMDAM_BOM_DATA_ENTRY_1000_SALES_PRICE_DETERMINATION_SEQ,False)
Call SetTextbox("Purchase Price Determ\. Seq\.","P_EKERV","",DT_ZMDAM_BOM_DATA_ENTRY_1000_PURCHASE_PRICE_DETERM_SEQ,False)
Call SetTextbox("List Variant","P_LIVAR","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LIST_VARIANT,False)
Call SetTextbox("List Group","P_LIGRU","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LIST_GROUP,False)

Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE_OCC1,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_ZMDAM_BOM_DATA_ENTRY_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_ZMDAM_BOM_DATA_ENTRY_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_ZMDAM_BOM_DATA_ENTRY_1000_VALIDITY),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_ZMDAM_BOM_DATA_ENTRY_1000_TO),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetGridContent("","","Final price", 1,"Article",DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE_OCC1,"DT_ZMDAM_BOM_GETCELLVALUE_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ZMDAM_BOM_GETCELLVALUE_OUTPUT",DT_ZMDAM_BOM_DATA_ENTRY_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetGridData("",1,"Final price",DT_ENTRY_GRIDCELL_FINAL_PRICE,False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("item1","DT_PRICING_DOC_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_PRICING_DOC_NUMBER_OUTPUT",DT_PRICING_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_ZMDAM_BOM_DATA_ENTRY_1000_CHECK_TEXT_OF_STATUSBAR_OCC1)


''''--------TransactionCode-VKP5----------''''

Call SetTcode(DT_ZMDAM_BOM_DATA_ENTRY_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_ZMDAM_BOM_DATA_ENTRY_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE,True)
Call ClickButton("Check entries   \(Enter\)",True)
Call ClickButton("Copy   \(F8\)",True)

Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_ZMDAM_BOM_DATA_ENTRY_1000_SALES_ORGANIZATION_OCC1,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_ZMDAM_BOM_DATA_ENTRY_1000_DISTRIBUTION_CHANNEL_OCC1,False)
Call SetTextbox("Validity Period","S_DATUM-LOW","",ConvertDate(DT_ZMDAM_BOM_DATA_ENTRY_1000_VALIDITY_PERIOD),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_ZMDAM_BOM_DATA_ENTRY_1000_TO_OCC1),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetGridContent("Sales Price Conditions","","Condition rate",1,"Index",1,"DT_GETCELLVALUE_OF_GRIDCELL_0_KBETR_OUTPUT")
Call GetGridContent("Sales Price Conditions","","Condition rate",2,"Index",2,"DT_GETCELLVALUE_OF_GRIDCELL_1_KBETR_OUTPUT")
Call GetGridContent("Sales Price Conditions","","Condition rate",3,"Index",3,"DT_GETCELLVALUE_OF_GRIDCELL_2_KBETR_OUTPUT")
Call GetGridContent("Sales Price Conditions","","Condition rate",4,"Index",4,"DT_GETCELLVALUE_OF_GRIDCELL_3_KBETR_OUTPUT")

Call VerifyGridCellContent("Sales Price Conditions",3,"Condition rate","",DT_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KBETR_OCC1)
Call VerifyGridCellContent("Sales Price Conditions",4,"Condition rate","",DT_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KBETR_OCC2)

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



