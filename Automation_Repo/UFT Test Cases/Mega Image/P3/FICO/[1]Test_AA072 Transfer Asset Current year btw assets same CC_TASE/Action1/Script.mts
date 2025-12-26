'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AA072 Transfer Asset Current year btw assets same CC_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 24th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AA072 same CC_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_AA072 Transfer Asset Current year btw assets same CC_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'reload DS to update dates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DESC_INC",(Cint(DT_DESC_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Asset values   \(Ctrl\+F1\)",False) 
Call ClickButtonIfExist("Continue   \(Enter\)",True) 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Fiscal year","EDIT_JAHRE","",DT_AS03_0100_FISCAL_YEAR,False)

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectNodeGuiTree("0","Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

Call VerifyGridCellContent("Transactions",1,"Asset Value Date","",DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT)
Call GetGridContentByTitle("Transactions","","Amount posted",1,"DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectTab("IDC_TABSTRIP",DT_AS03_0100_POSTED_VALUES,False)
'Capture the screenshot
Call TakeScreenShot()

Call GetGridContentByTitle("Posted dep\. values Local GAAP APC, depreciation","","Posted values",11,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OUTPUT")
Call GetGridContentByTitle("Posted dep\. values Local GAAP APC, depreciation","","Posted values",5,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",1,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",2,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFAZ_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("Posted dep\. values Local GAAP APC, depreciation",5,"Posted values","",DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE)

Call SelectRowGuiGridbyRowNo("Posted dep\. values Local GAAP APC, depreciation","",DT_AS03_0302_GRIDCELL_4_POSTED_VALUES,False)
Call SelectRowGuiGridbyRowNo("Depreciation posted/planned","",DT_AS03_0302_GRIDCELL_0_ORD_DEP_OCC1,False)
Call SelectNodeGuiTree("0","Depreciation Areas;0L Leading Ledger")
'Capture the screenshot
Call TakeScreenShot()

Call GetGridContentByTitle("Posted dep\. values Local GAAP APC, depreciation","","Posted values",4,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OUTPUT")
Call GetGridContentByTitle("Posted dep\. values Local GAAP APC, depreciation","","Posted values",5,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE_OCC1_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",1,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",2,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",3,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Posted dep\. values Local GAAP APC, depreciation","","Posted values",11,"DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("Posted dep\. values Local GAAP APC, depreciation",5,"Posted values","",DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Class","ANLA-ANLKL","",DT_AS03_1000_CHECK_TEXT_OF_CLASS,False)
'
'''----------------------Tcode AS01----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS03_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS03_1000_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS03_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS03_0105_NUMBER_OF_SIMILAR_ASSETS,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS03_1140_DESCRIPTION,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS03_1145_COST_CENTER,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS03_1145_BUSINESS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS03_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS03_1160_EVALUATION_GROUP_1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Origin",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS03_1181_VENDOR,False)
Call PressEnter() 
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS03_1182_WBS_ELEMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Leasing",False)
'Capture the screenshot
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","Deprec. Areas",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLAISTTC_ANLB","UseLife",1,DT_AS03_1190_TABLECELL_USELIFE_0,False)
Call SetTableDataNoRef("SAPLAISTTC_ANLB","Prd",1,DT_AS03_1190_TABLECELL_PRD_0,False)

Call SetTableDataNoRef("SAPLAISTTC_ANLB","UseLife",2,DT_AS03_1190_TABLECELL_USELIFE_1,False)
Call SetTableDataNoRef("SAPLAISTTC_ANLB","Prd",2,DT_AS03_1190_TABLECELL_PRD_1,False)

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call PressEnter() 
Call PressEnter()  
Call PressEnter()  
Call PressEnter()  
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyCheckBoxValue("ANLZ-XSTIL",DT_AS03_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN)
Call SelectCheckbox("ANLZ-XSTIL","1",DT_AS03_1145_ASSET_SHUTDOWN,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item2","DT_AS03_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetStatusBar("item1","DT_AS03_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS03_0105_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

