

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
'.................Test Script Name : Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P5
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P5"
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

'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)


'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-NDJAR",0,"DT_USEFULLIFE_OUTPUT",False)

'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)
DT_GRIDTITLE = "IFRS APC, depreciation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""

Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,8,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_7_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,9,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_8_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,10,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_9_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,11,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_10_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,12,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_11_NAFA)


Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")


'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-NDJAR",0,"DT_USEFULLIFE_OUTPUT1",False)

'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)

DT_GRIDTITLE = "Local GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""
Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC1)


Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;20 Pre-merger STA valuation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-NDJAR",0,"DT_USEFULLIFE_OUTPUT2",False)

'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)
DT_GRIDTITLE = "Pre-merger STA valuation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""

Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC4)
Call VerifyGridCellContent(DT_GRIDTITLE,3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,8,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,9,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,10,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,11,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
'Call VerifyGridCellContent(DT_GRIDTITLE,12,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_11_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,12,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_11_NAFA_OCC3)
Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")


'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call GetTextboxValue("AW01_DEP_PAR-NDJAR",0,"DT_USEFULLIFE_OUTPUT3",False)

'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)

DT_GRIDTITLE = "IFRS LOCL GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""

'Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC4)
Call VerifyGridCellContent(DT_GRIDTITLE,3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,8,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,9,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,10,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
Call VerifyGridCellContent(DT_GRIDTITLE,11,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
'Call VerifyGridCellContent(DT_GRIDTITLE,12,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_11_NAFA)
Call VerifyGridCellContent(DT_GRIDTITLE,12,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_11_NAFA_OCC3)

''----------------------Tcode AS02----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS03_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS03_0100_OKCD)
Call PressEnter() 


'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)


Call SetTableData("SAPLAISTTC_ANLB","DKey","3","","",DT_AS03_1190_TABLECELL_DKEY_2,False)
Call SetTableData("SAPLAISTTC_ANLB","UseLife","1","","",DT_AS03_1190_TABLECELL_USELIFE_0,False) 
Call SetTableData("SAPLAISTTC_ANLB","UseLife","3","","",DT_AS03_1190_TABLECELL_USELIFE_2,False) 
Call SetTableData("SAPLAISTTC_ANLB","UseLife","4","","",DT_AS03_1190_TABLECELL_USELIFE_0,False) 
Call SetTableData("SAPLAISTTC_ANLB","UseLife","5","","",DT_AS03_1190_TABLECELL_USELIFE_0,False) 
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
Call GetStatusBar("item1","DT_ASSET_NO_OUTPUT")
Call VerifyStatusBar("The asset "&DT_ASSET_NO_OUTPUT&" 0 is changed" )


'----------------------Tcode AS03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS03_0100_OKCD_OCC1) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS03_0100_OKCD_OCC1)
Call PressEnter() 

'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)
Call TakeScreenShot()

Call VerifyTableCellContent(3,"DKey","SAPLAISTTC_ANLB",DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_DKEY_2)
Call VerifyTableCellContent(1,"UseLife","SAPLAISTTC_ANLB",DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(3,"UseLife","SAPLAISTTC_ANLB",DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2)
Call VerifyTableCellContent(4,"UseLife","SAPLAISTTC_ANLB",DT_AS03_0304_CHECK_TEXT_OF_USEFUL_LIFE_OCC2)
Call VerifyTableCellContent(5,"UseLife","SAPLAISTTC_ANLB",DT_AS03_0304_CHECK_TEXT_OF_USEFUL_LIFE_OCC2)

'
'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0_OCC1,False)

'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)
Wait(2)
''
''
DT_GRIDTITLE = "IFRS APC, depreciation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""

'''Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC2)
''Call VerifyGridCellContent(DT_GRIDTITLE,3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFA_OCC1)
''Call VerifyGridCellContent(DT_GRIDTITLE,4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFA_OCC1)
''Call VerifyGridCellContent(DT_GRIDTITLE,5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFA_OCC1)
''Call VerifyGridCellContent(DT_GRIDTITLE,6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_NAFA_OCC1)
''Call VerifyGridCellContent(DT_GRIDTITLE,7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_NAFA_OCC1)

Call VerifyGridCellContent("IFRS.*",3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFA_OCC1)
Call VerifyGridCellContent("IFRS.*",4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFA_OCC1)
Call VerifyGridCellContent("IFRS.*",5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFA_OCC1)
Call VerifyGridCellContent("IFRS.*",6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_NAFA_OCC1)
Call VerifyGridCellContent("IFRS.*",7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_NAFA_OCC1)

Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")


'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS03_0304_CHECK_TEXT_OF_USEFUL_LIFE_OCC3,False)

'Navigate to Comparisons Tab

'Call VerifyGridCellContent("Local GAAP APC, depreciation:2021 -2026",3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFA_OCC2)
Call SelectTab("IDC_TABSTRIP","Comparisons",False)

'DT_GRIDTITLE = "LOCAL GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+Cint(5)&""

Call VerifyGridCellContent("LOCAL GAAP APC.*",2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFA_OCC3)
Call VerifyGridCellContent("LOCAL GAAP APC.*",3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFA_OCC2)
Call VerifyGridCellContent("LOCAL GAAP APC.*",4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFA_OCC2)
Call VerifyGridCellContent("LOCAL GAAP APC.*",5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFA_OCC2)
Call VerifyGridCellContent("LOCAL GAAP APC.*",6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_5_NAFA_OCC2)
Call VerifyGridCellContent("LOCAL GAAP APC.*",7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_6_NAFA_OCC2)


Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;20 Pre-merger STA valuation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

''Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3,False)
Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0_OCC1,False)
'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)
''DT_GRIDTITLE = "Pre-merger STA valuation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""

''Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_1)
''Call VerifyGridCellContent(DT_GRIDTITLE,3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_2)
''Call VerifyGridCellContent(DT_GRIDTITLE,4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_3)
''Call VerifyGridCellContent(DT_GRIDTITLE,5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_4)
''Call VerifyGridCellContent(DT_GRIDTITLE,6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_5)
''Call VerifyGridCellContent(DT_GRIDTITLE,7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_6)
Call VerifyGridCellContent("PRE.*",2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_1)
Call VerifyGridCellContent("PRE.*",3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_2)
Call VerifyGridCellContent("PRE.*",4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_3)
Call VerifyGridCellContent("PRE.*",5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_4)
Call VerifyGridCellContent("PRE.*",6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_5)
Call VerifyGridCellContent("PRE.*",7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_20_6)

Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")


'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

''Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_4,False)
Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_AS03_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0_OCC1,False)
'Navigate to Comparisons Tab
Call SelectTab("IDC_TABSTRIP","Comparisons",False)

''DT_GRIDTITLE = "IFRS LOCL GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+Cint(DT_USEFULLIFE_OUTPUT)&""
''Call VerifyGridCellContent(DT_GRIDTITLE,2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_1)
''Call VerifyGridCellContent(DT_GRIDTITLE,3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_2)
''Call VerifyGridCellContent(DT_GRIDTITLE,4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_3)
''Call VerifyGridCellContent(DT_GRIDTITLE,5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_4)
''Call VerifyGridCellContent(DT_GRIDTITLE,6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_5)
''Call VerifyGridCellContent(DT_GRIDTITLE,7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_6)
Call VerifyGridCellContent("IFRS.*",2,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_1)
Call VerifyGridCellContent("IFRS.*",3,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_2)
Call VerifyGridCellContent("IFRS.*",4,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_3)
Call VerifyGridCellContent("IFRS.*",5,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_4)
Call VerifyGridCellContent("IFRS.*",6,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_5)
Call VerifyGridCellContent("IFRS.*",7,"Ordinary deprec.",0,DT_AS03_0303_CHECK_22_6)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

