
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P6
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P6
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P6"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode AS03----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 

Call SelectTab("TABSTRIP100","Deprec. Areas",False)
Call TakeScreenShot()
'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-AFABG",0,"DT_ORD_DEP_START_DATE",False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)

'Call GetGridContent("Depreciation posted/planned",0,"Depreciation period",2,"","","DT_Depreciation_Period")
'Call GetGridContent("Depreciation posted/planned",0,"Ordinary dep. TBP",2,"","","DT_ORDINARY_DEP_PERIOD")

Call GetGridContentByTitle("Depreciation posted/planned",0,"Depreciation period",2,"DT_Depreciation_Period")
Call GetGridContentByTitle("Depreciation posted/planned",0,"Ordinary dep. TBP",2,"DT_ORDINARY_DEP_PERIOD")

Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-AFABG",0,"DT_ORD_DEP_START_DATE1",False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)

''Call GetGridContent("Depreciation posted/planned",0,"Depreciation period",3,"","","DT_Depreciation_Period1")
''Call GetGridContent("Depreciation posted/planned",0,"Ordinary dep. TBP",3,"","","DT_ORDINARY_DEP_PERIOD1")

Call GetGridContentByTitle("Depreciation posted/planned",0,"Depreciation period",2,"DT_Depreciation_Period_OCC1")
Call GetGridContentByTitle("Depreciation posted/planned",0,"Ordinary dep. TBP",2,"DT_ORDINARY_DEP_PERIOD_OCC1")


Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;20 Pre-merger STA valuation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-AFABG",0,"DT_ORD_DEP_START_DATE2",False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)

''Call GetGridContent("Depreciation posted/planned",0,"Depreciation period",3,"","","DT_Depreciation_Period2")
''Call GetGridContent("Depreciation posted/planned",0,"Ordinary dep. TBP",3,"","","DT_ORDINARY_DEP_PERIOD2")

Call GetGridContentByTitle("Depreciation posted/planned",0,"Depreciation period",2,"DT_Depreciation_Period_OCC2")
Call GetGridContentByTitle("Depreciation posted/planned",0,"Ordinary dep. TBP",2,"DT_ORDINARY_DEP_PERIOD_OCC2")

Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-AFABG",0,"DT_ORD_DEP_START_DATE3",False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)

''Call GetGridContent("Depreciation posted/planned",0,"Depreciation period",3,"","","DT_Depreciation_Period3")
''Call GetGridContent("Depreciation posted/planned",0,"Ordinary dep. TBP",3,"","","DT_ORDINARY_DEP_PERIOD3")

Call GetGridContentByTitle("Depreciation posted/planned",0,"Depreciation period",2,"DT_Depreciation_Period_OCC3")
Call GetGridContentByTitle("Depreciation posted/planned",0,"Ordinary dep. TBP",2,"DT_ORDINARY_DEP_PERIOD_OCC3")


''----------------------Tcode AS02----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS03_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS03_0100_OKCD)
Call PressEnter() 


'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)

Call SetTableData("SAPLAISTTC_ANLB","ODep Start","1","","",ConvertDate(DT_AS03_1190_TABLECELL_ODEP_START_0),False) 
Call SetTableData("SAPLAISTTC_ANLB","ODep Start","3","","",ConvertDate(DT_AS03_1190_TABLECELL_ODEP_START_2),False) 
Call SetTableData("SAPLAISTTC_ANLB","ODep Start","4","","",ConvertDate(DT_AS03_1190_TABLECELL_ODEP_START_2),False) 
Call SetTableData("SAPLAISTTC_ANLB","ODep Start","5","","",ConvertDate(DT_AS03_1190_TABLECELL_ODEP_START_2),False) '

Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()

'Post the Document
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call PressEnter() 
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call TakeScreenShot()
Wait 2

Call GetStatusBar("item1","DT_ASSET_NO_OUTPUT")
'Call VerifyStatusBarMessageType("S")
Call PressEnter()
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode AS03----------------------------

Call SetTcode(DT_AS03_0100_OKCD_OCC1) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS03_0100_OKCD_OCC1)
Call PressEnter() 
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)

Call VerifyTableCellContent(1,"ODep Start","SAPLAISTTC_ANLB",ConvertDAte(DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_0))
Call VerifyTableCellContent(3,"ODep Start","SAPLAISTTC_ANLB",ConvertDAte(DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_2))
Call VerifyTableCellContent(4,"ODep Start","SAPLAISTTC_ANLB",ConvertDAte(DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_2))
Call VerifyTableCellContent(5,"ODep Start","SAPLAISTTC_ANLB",ConvertDAte(DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_ODEP_START_2))

'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG",0,ConvertDAte(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC2),False)
''Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG",0,Replace(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC2,"/","."),False)
'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call TakeScreenShot()

'
Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG",0,ConvertDAte(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC3),False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call TakeScreenShot()

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;20 Pre-merger STA valuation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG",0,ConvertDAte(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC3),False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call TakeScreenShot()

Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG",0,ConvertDAte(DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE_OCC3),False)

'Navigate to Posted Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call TakeScreenShot()


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
 

