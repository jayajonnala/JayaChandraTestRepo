

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P5
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

gstrTestCaseName = "Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P7"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'----------------------Tcode AS02----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 


'Navigate to Time-dependent Tab
Call SelectTab("TABSTRIP100","Time-dependent",False)
Call TakeScreenShot()
Call Clickbutton("More Intervals",False)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call PressEnter() 
Call TakeScreenShot()

Call SetTableData("SAPLAISTTIME","Shutdown","2","","",DT_AS02_3000_TABLECELL_SHUTDOWN_1,False)
Call SetTableData("SAPLAISTTIME","Shutdown","1","","",DT_AS02_3000_TABLECELL_SHUTDOWN_0,False)

Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call Clickbutton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call PressEnter() 
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call PressEnter() 
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(Lcase(DT_AS02_0100_CHECK_TEXT_OF_STATUSBAR))

'----------------------Tcode AS03----------------------------

'''''Create Purchase Order
Call SetTcode(DT_AS02_0100_OKCD) 
Call PressEnter()
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0,"#1;#1;#1")
''Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call FindRowNumber("Depreciation posted/planned", "Depreciation period", "12", "DT_ROW_OUTPUT")
Call VerifyGridCellContent("Depreciation posted/planned",DT_ROW_OUTPUT,"NAFAZ",0,DT_AS02_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ)

Call ActivateNodeGuiTree(0,"#1;#2;#1")
Call FindRowNumber("Depreciation posted/planned", "Depreciation period", "12", "DT_ROW_OUTPUT")
Call VerifyGridCellContent("Depreciation posted/planned",DT_ROW_OUTPUT,"NAFAZ",0,DT_AS02_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ_OCC1)

Call ActivateNodeGuiTree(0,"#1;#1;#2")
Call FindRowNumber("Depreciation posted/planned", "Depreciation period", "12", "DT_ROW_OUTPUT")
Call VerifyGridCellContent("Depreciation posted/planned",DT_ROW_OUTPUT,"NAFAZ",0,DT_AS02_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_8_NAFAZ_OCC1)

Call ActivateNodeGuiTree(0,"#1;#3;#1")
Call FindRowNumber("Depreciation posted/planned", "Depreciation period", "12", "DT_ROW_OUTPUT")
Call VerifyGridCellContent("Depreciation posted/planned",DT_ROW_OUTPUT,"NAFAZ",0,DT_AS02_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_8_NAFAZ_OCC1)

''Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)

'DT_GRIDNAME = "IFRS LOCL GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+5&""
'DT_GRIDNAME = "IFRS LOCL GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+35&""
DT_GRIDNAME = "IFRS LOCL GAAP APC, depreciation:.*"

Call VerifyGridCellContent(DT_GRIDNAME,2,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,3,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,4,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,5,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,6,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,7,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_AHK)

Call ActivateNodeGuiTree(0,"#1;#1;#1")
'DT_GRIDNAME = "IFRS APC, depreciation:"&Year(Date)&" -"&Year(Date)+5&""
'DT_GRIDNAME = "IFRS APC, depreciation:"&Year(Date)&" -"&Year(Date)+35&""
DT_GRIDNAME = "IFRS APC, depreciation:.*"
Call VerifyGridCellContent(DT_GRIDNAME,2,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,3,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,4,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,5,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,6,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,7,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_AHK)


Call ActivateNodeGuiTree(0,"#1;#2;#1")
'DT_GRIDNAME = "Local GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+5&""
'DT_GRIDNAME = "Local GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+35&""
DT_GRIDNAME = "Local GAAP APC, depreciation:.*"

Call VerifyGridCellContent(DT_GRIDNAME,2,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,3,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,4,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,5,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,6,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,7,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_AHK)

Call ActivateNodeGuiTree(0,"#1;#1;#2")
'DT_GRIDNAME = "Pre-merger STA valuation:"&Year(Date)&" -"&Year(Date)+5&""
'DT_GRIDNAME = "Pre-merger STA valuation:"&Year(Date)&" -"&Year(Date)+35&""
DT_GRIDNAME = "Pre-merger STA valuation:.*"
Call VerifyGridCellContent(DT_GRIDNAME,2,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,3,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,4,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,5,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,6,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AHK)
Call VerifyGridCellContent(DT_GRIDNAME,7,"Acquisition value",0,DT_AS02_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_AHK)

Call LogOff()
Call FinalStatus()
