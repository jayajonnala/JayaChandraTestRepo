

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P3
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''''----------------------Tcode AS03----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot    

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call TakeScreenShot   

Call ClickBUtton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot   
Call ActivateNodeGuiTree(0, "#1;#1;#1")
Call TakeScreenShot   
''Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
''Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
''Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
''Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, "")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, "")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, "")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, "")

Call ActivateNodeGuiTree(0, "#1;#2;#1")
Call TakeScreenShot   
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))

Call SelectTab("IDC_TABSTRIP","Parameters",False)
Call TakeScreenShot   
Call VerifyTextBoxContent("Ord\.dep\.start date","AW01_DEP_PAR-AFABG","",DT_AS03_0304_CHECK_TEXT_OF_ORDDEPSTART_DATE,False)

Call SelectTab("IDC_TABSTRIP","Comparisons",False)
Call TakeScreenShot  
Call SetTextBox("To","EDIT_JAHRE_TO","",DT_AS03_0202_TO,False)
Call PressEnter()     
Call TakeScreenShot 

Call GetNumberOfRows("shell",0,"DT_NUMBER_ROWS")

Call verifyGridRowCount("shell",0,DT_NUMBER_ROWS)
Call VerifyGridCellContent("Local GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+10&"", 2, "BEW", "",DT_AS03_0303_CHECK_GETCELLVALUE_OF_GRIDCELL_12_BEW)
Call VerifyGridCellContent("Local GAAP APC, depreciation:"&Year(Date)&" -"&Year(Date)+10&"", 13, "BEW", "",DT_ORDINARY_DEPREC)

Call SelectTab("IDC_TABSTRIP","Posted values",False)
Call GetNumberOfRows("shell",0,"DT_NUMBER_ROWS")
Call TakeScreenShot 
Call VerifyGridCellContentByRefColumn("Depreciation posted/planned","","Depreciation period","11","Ordinary dep. TBP",0,DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFAZ)
''Call VerifyGridCellContent("Depreciation posted/planned", 3, "NAFAZ", "",DT_AS03_0302_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NAFAZ)

Call ClickButton("Exit   \(Shift\+F3\)",False)

Call ClickButton("Exit   \(Shift\+F3\)",False)

''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************




