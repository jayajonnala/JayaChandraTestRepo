

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Article Creation via DAP_v3_Sticker management
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


gstrTestCaseName = "Test_Article Creation_v3_Stickert"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
''''''''''''''''''''-----------------------------------
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''''----------------------Tcode  ZMDAM_BOM_DATA_ENTRY ----------------------------
''''Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Article","P_MATNR","",DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE,False)
Call SetTextbox("Link Type","P_STLAN","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LINK_TYPE,False)
Call SetTextbox("Link retention level","P_STLAL","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LINK_RETENTION_LEVEL,False)
Call SetTextbox("Layout","P_VARI","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LAYOUT,False)


'Click on Execute
Call ClickButton("Execute   \(F8\)",False)
Wait 5
Call TakeScreenShot()

Call ClickButtonToolBar("CHANGE", 0)
Call SetTextbox("BOM Group","GV_LABOR","",DT_BOM_STICKER_TYPE,False)
Wait 5

SAPGuiSession("Session").SAPGuiWindow("DG: BOM Update").InsightObject("InsightObject_2").Click
'SAPGuiSession("Session").SAPGuiWindow("DG: BOM Update").InsightObject("InsightObject").Click
Wait 5
Call SetGridData("Change/Add Item to BOM", 1, "MATNR", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_PARENT_ARTICLE, False)
Call SetGridData("Change/Add Item to BOM", 1, "MENGE", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_QUANTITY, False)
Call DoubleClickGuiGridCell("Change/Add Item to BOM", 0, 1, "MENGE", False)
Wait 2
Call TakeScreenShot()
Call SetGridData("Change/Add Item to BOM", 1, "NO_INDX", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_MAIN_PARENT, False)
Call SetGridData("Change/Add Item to BOM", 1, "AUSCH", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_FREE_QUANTITY, False)
Call DoubleClickGuiGridCell("Change/Add Item to BOM", 0, 1, "AUSCH", False)
Wait 2
Call TakeScreenShot()
SAPGuiSession("Session").SAPGuiWindow("DG: BOM Update").InsightObject("InsightObject_2").Click
'SAPGuiSession("Session").SAPGuiWindow("DG: BOM Update").InsightObject("InsightObject").Click
Wait 5
Call SetGridData("Change/Add Item to BOM", 1, "MATNR", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_1_PARENT_ARTICLE, False)
Call SetGridData("Change/Add Item to BOM", 1, "MENGE", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_1_QUANTITY, False)
Call DoubleClickGuiGridCell("Change/Add Item to BOM", 0, 1, "MENGE", False)
Wait 2
Call TakeScreenShot()
Call SetGridData("Change/Add Item to BOM", 1, "AVOAU", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_1_FREE_CONTENT, False)
Call DoubleClickGuiGridCell("Change/Add Item to BOM", 0, 1, "AVOAU", False)
Wait 2
Call TakeScreenShot()
Call ClickButtonToolBar("SAVE",0)
Call TakeScreenShot()
Call ClickButtonIfExist("Yes",True)
Call GetLabelContentByRefLabel("Message text", 0, -32, "DT_ZMDAM_BOM_DATA_ENTRY_0120_CHECK_TEXT_OF_BOMGROUP_00043798S_HAS_BEEN_CREATED_EXTER_OUTPUT", True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyifGuiLabelExists(DT_ZMDAM_BOM_DATA_ENTRY_0120_CHECK_TEXT_OF_BOMGROUP_00043798S_HAS_BEEN_CREATED_EXTER_OUTPUT)
Call TakeScreenShot()
Call ClickButton("Copy   \(Enter\)",True)
Wait 5
Call ClickButtonifexist("Continue   \(Enter\)",True)
Call TakeScreenShot()

CAll SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY", "Calc. Type", 1, "%", False)
CAll SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, DT_ZMDAM_BOM_DATA_ENTRY_1960_TABLECELL_AMOUNT_0, False)
Call TakeScreenShot()
Call PressEnter()

Call SelectRowGuiTableByRow("SAPMV13ATCTRL_FAST_ENTRY", 1, False)
Call ClickButton("Save   \(Ctrl\+S\)",Fale)
Call TakeScreenShot()

Call VerifyStatusBar(DT_ZMDAM_BOM_DATA_ENTRY_1000_CHECK_TEXT_OF_STATUSBAR)

'''''''----------------------Tcode  VKP5 ----------------------------
'''''Enter the transaction code
Call SetTcode(DT_ZMDAM_BOM_DATA_ENTRY_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Created By","ENAME-LOW","","",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Wait 2
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RWVKP007", 0, 6, True)
Call takeScreenShot()
Wait 2
Call ClickButtonIfExist("Choose   \(F2\)",True)
Wait 5
Call Clickbutton("%_S_MATNR_%_APP_%-VALU_PUSH",False)

Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_ZMDAM_BOM_DATA_ENTRY_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_ZMDAM_BOM_DATA_ENTRY_3010_TABLECELL_SINGLE_VALUE_1,True)
Call takeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call takeScreenShot()

Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_ZMDAM_BOM_DATA_ENTRY_1000_TO),False)

Call ClickButton("Execute   \(F8\)",FalsE)
Call takeScreenShot()
Call ClickButtonToolBar("&REFRESH", 0)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_ZMDAM_BOM_DATA_ENTRY_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter()

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

