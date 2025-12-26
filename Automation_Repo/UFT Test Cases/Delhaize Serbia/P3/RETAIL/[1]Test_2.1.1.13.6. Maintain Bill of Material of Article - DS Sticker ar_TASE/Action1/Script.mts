

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.1.1.13.6. Maintain Bill of Material of Article - DS Sticker ar
'.................Author : TCS 	   :
'................ Creation Date    : 
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

gstrTestCaseName = "Test_2.1.1.13.6. Maintain Bill of Material of Article - DS Sticker ar"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.1.1. Create Article Purchasing Data.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=7
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''----------------------Tcode-ZMDAM_BOM_DATA_ENTRY ----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Article","P_MATNR","",DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE,False)
Call SetTextbox("Link Type","P_STLAN","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LINK_TYPE,False)
Call SetTextbox("Link retention level","P_STLAL","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LINK_RETENTION_LEVEL,False)
Call SetTextbox("Layout","P_VARI","",DT_ZMDAM_BOM_DATA_ENTRY_1000_LAYOUT,False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call ClickButtonToolBar("CHANGE", 0)
Wait 2
SAPGuiSession("sessionObject").SAPGuiWindow("objWindow").InsightObject("InsightObject").Click

Call SetGridData("Change/Add Item to BOM", 1, "MATNR", " ", False)
Wait 2
Call SetGridData("Change/Add Item to BOM", 1, "MATNR", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_PARENT_ARTICLE_OCC1, False)
Call SetGridData("Change/Add Item to BOM", 1, "MENGE", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_QUANTITY, False)
Call SetGridData("Change/Add Item to BOM", 1, "MEINS", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_UNIT_OF_MEASURE_PARENT, False)
Call SetGridData("Change/Add Item to BOM", 1, "NO_INDX", DT_ZMDAM_BOM_DATA_ENTRY_0100_GRIDCELL_0_MAIN_PARENT, False)
Call TakeScreenShot()

Call SelectRowGuiGridbyRowNo("Change/Add Item to BOM", 0, 1, False)
Wait 2
Call ClickButtonToolBar("SAVE", 0)
Call TakeScreenShot()

Call ClickButtonifExist("YES",True)
Call TakeScreenShot()

Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot()
Wait 2
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC2) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Condition Type","RV13A-KSCHL","","ZSTI",False)
Call PressEnter()
Call TakeScreenShot()

Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Article", 1, "", "", DT_ZMDAM_BOM_DATA_ENTRY_1000_ARTICLE, False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Calc. Type", 1, "", "", "%", False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY", "Amount", 1, "", "", DT_ZMDAM_BOM_DATA_ENTRY_1960_TABLECELL_AMOUNT_0, False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_ZMDAM_BOM_DATA_ENTRY_1000_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()

