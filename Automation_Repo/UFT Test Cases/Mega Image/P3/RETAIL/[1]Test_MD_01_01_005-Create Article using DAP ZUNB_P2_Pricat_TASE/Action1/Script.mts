

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_005-Create Article using DAP ZUNB_P2_Pricat
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

gstrTestCaseName = "Test_MD_01_01_005- DAP ZUNB_P2_Pricat"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode SBWP----------------------------

'Enter the Transaction Code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Activate the Node
Call ActivateNodeGuiTree(0,"#1;#1;#3")
Wait(2)
Call TakeScreenShot()

'Click on Filtr button
Call ClickButtonToolBar("&MB_FILTER",0)
Call TakeScreenShot()
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot()
'Search the Filter item
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Work Item Title",True) 
Call TakeScreenShot()

'Click on OK
Call ClickButton("OK   \(Enter\)",True) 
Wait(2)

'Verify If Hit Is found
Call VerifyTextBoxContent("Information Message","GS_SEARCH-SEARCH_INFO",0,Lcase("Hit displayed : 1"),True)


'Click on OK
Call ClickButton("Cancel   \(F12\)",True) 
Wait(2)


Call ClickButton("Add Filter Criterion \(F7\)",True) 
Wait(2)

Call ClickButton("600_BUTTON",True) 
Wait(2)

Call SetTextbox("Work Item Title","%%DYN001-LOW","",DT_SBWP_1105_WORK_ITEM_TITLE,True) 
Call TakeScreenShot()

Call ClickButton("Execute   \(Enter\)",True) 

'Approval/Enrichment for Article: AUTO_EN_TEST216 required 
Call SelectRowGuiGrid("Workflow.*",0,"Title",DT_SBWP_1000_GRIDCELL_0_TITLE,False)
Call DoubleClick()
Wait(1)
Call TakeScreenShot()
Call ClickButton("Submit completed Form to PRICAT",False) 
Wait(2)
Call TakeScreenShot()

''----------------------Tcode W_SYNC ----------------------------
'Enter the transaction code
Call SetTcode(DT_SBWP_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenSHot()

Call SelectCheckbox("P_MYITEM",0,"OFF",False)
Call SelectCheckbox("P_ONLY",0,"OFF",False)
Call SelectCheckbox("P_COMBOM",0,"OFF",False)
Call SelectCheckbox("P_COMVAR",0,"OFF",False)

Call SetTextbox("Profile","P_PROFIL","",DT_SBWP_0500_PROFILE,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("GTIN Base UoM","PI_06-LOW","",DT_SBWP_0500_GTIN_BASE_UOM,False)

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()
Wait 10
'Click on Select Button
Call ClickButton("Select   \(F8\)",False)

'Click on Select Button
Call ClickButton("Select   \(F8\)",False)

'Click on Select Button
Call ClickButton("Select   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ActivateNodeGuiTree(0,"#1;#1;#1")
Wait(2)
Call TakeScreenShot()

'Verify the grid content
Call VerifyGridCellContent("",1,"GTIN Base UoM",0,DT_SBWP_5000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GTIN)

'Click on Link to Price Catalog Maintenance
Call SelectRowGuiGrid("",0,"GTIN Base UoM",DT_SBWP_5000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GTIN,False)
Call ClickButton("Link to Price Catalog Maintenance   \(F7\)",False)

'Select Row
Call SelectRowGuiGrid("",0,"Processing Status","01",False)
Wait(1)

'Click on Create/Change
Call ClickButton("Create/Change Article - Immediately   \(Shift\+F9\)",False)
Wait(10)

Call SelectRowGuiGrid("",0,"Processing Status","03",False)
Wait(1)

'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
'Refresh the screen
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)
Call ClickButtonToolBar("&REFRESH",0)
Wait(5)

Call GetGridContentByRefColumn("",0,"Processing Status","04","Article","DT_SBWP_0400_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MATNR_OUTPUT")
Wait 2
Call FindRowNumber("", "Processing Status", "04", "DT_ROWNUMBER")
Call ClickCellGuiGrid("", 0, "Article", DT_ROWNUMBER, "", "", False)
Call TakeScreenShot()

Call SetTextBox("Purchasing Org\.","RMMW1-EKORG",0,DT_SBWP_0100_PURCHASING_ORG,False)
Call PressEnter()
Call TakeScreenShot()

Call VerifyTextBoxContent("Single article","MAKT-MAKTX", 0, Lcase(DT_SBWP_1010_CHECK_TEXT_OF_SINGLE_ARTICLE), False)
Call VerifyTextBoxContent("Material Type","MARA-MTART", 0, DT_SBWP_2001_CHECK_TEXT_OF_ARTICLE_TYPE, False)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************







