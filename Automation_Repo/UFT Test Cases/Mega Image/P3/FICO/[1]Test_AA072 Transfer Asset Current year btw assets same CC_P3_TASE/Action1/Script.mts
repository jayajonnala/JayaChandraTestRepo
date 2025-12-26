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
'.................Test Script Name : Test_AA072 Transfer Asset Current year btw assets same CC_P3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 24th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AA072 assets same CC_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_AA072 Transfer Asset Current year btw assets same CC_P3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''reload
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

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0100_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0100_COMPANY_CODE,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABUMN_0100_SUBNUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Asset values   \(Ctrl\+F1\)",False) 
Call ClickButtonIfExist("Continue   \(Enter\)",True) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectNodeGuiTree("0","Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","1",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","1","")
'Capture the screenshot
Call TakeScreenShot()



Call SelectTab("IDC_TABSTRIP",DT_ABUMN_0100_POSTED_VALUES,False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",4,"Change","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",4,"Posted values","","")

'get grid contents
Call GetGridContentByTitle("Posted dep\. values IFRS APC, depreciation","","Posted values",5,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",1,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",2,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",3,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NAFAZ_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",5,"Posted values","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE)
'
''
'''----------------------Tcode AS03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ABUMN_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ABUMN_0100_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0100_ASSET_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABUMN_0100_SUBNUMBER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Asset values   \(Ctrl\+F1\)",False) 
Call ClickButtonIfExist("Continue   \(Enter\)",True) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectNodeGuiTree("0","Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call GetGridContentByTitle("Planned values IFRS APC, depreciation","1","Year-end",11,"DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'validate details
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","1",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","1",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Year-end","1",DT_ABUMN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE)

'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("IDC_TABSTRIP",DT_ABUMN_0100_POSTED_VALUES_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()

'get grid contents
Call GetGridContentByTitle("Posted dep\. values IFRS APC, depreciation","","Posted values",11,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",5,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ_OUTPUT")
Call GetGridContentByTitle("Depreciation posted/planned","","Ordinary dep. TBP",3,"DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFAZ_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'verify details
Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",4,"Change","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",4,"Posted values","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",11,"Posted values","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE)
Call VerifyGridCellContent("Depreciation posted/planned",5,"Ordinary dep. TBP","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NAFAZ)

Call SelectRowGuiGridbyRowNo("Depreciation posted/planned","",DT_ABUMN_0302_GRIDCELL_4_ORD_DEP,False)
Call SelectNodeGuiTree("0","Depreciation Areas;0L Leading Ledger")
'Capture the screenshot
Call TakeScreenShot()
Call SelectTab("IDC_TABSTRIP",DT_ABUMN_0100_PLANNED_VALUES_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("IDC_TABSTRIP",DT_ABUMN_0100_POSTED_VALUES_OCC3,False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("Posted dep\. values IFRS APC, depreciation",11,"Posted values","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_10_JENDE_OCC1)
Call VerifyGridCellContent("Depreciation posted/planned",3,"Ordinary dep. TBP","",DT_ABUMN_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFAZ)



'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

