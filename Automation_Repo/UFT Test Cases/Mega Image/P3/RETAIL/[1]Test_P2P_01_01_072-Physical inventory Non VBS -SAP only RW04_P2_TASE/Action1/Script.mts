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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_072-Physical inventory Non VBS -SAP only RW04_P2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 9th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_072-Physical inventory P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_072-Physical inventory  Non VBS -SAP only  RW04_P2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
''''----------------------Tcode MI04----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Phys\. Inventory Doc\.","RM07I-IBLNR","",DT_MI04_0701_PHYS_INVENTORY_DOC_1,False)
Call SetTextbox("Fiscal Year","RM07I-GJAHR","",DT_MI04_0701_FISCAL_YEAR,False)
Call SetTextbox("Count Date","RM07I-ZLDAT","",DT_MI04_0701_COUNT_DATE,False)
Call PressEnter()   
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("ISEG-ERFMG","0",DT_MI04_0731_ISEGERFMG ,False)
Call SetTextboxNoLabel("ISEG-ERFMG","1",DT_MI04_0731_ISEGERFMG_OCC1 ,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTextStatusBar("DT_MI04_0701_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'verify
Call VerifyStatusBar(DT_MI04_0701_CHECK_TEXT_OF_STATUSBAR_OCC1)
''
''''----------------------Tcode MI24----------------------------
'Enter the Tcode
Call SetTcode(DT_MI04_0701_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MI04_0701_OKCD)
'Capture the screenshot
Call TakeScreenShot()

'multilpe selection for article
SAPGuiSession("transaction:=MI20").SAPGuiWindow("transaction:=MI20").SAPGuiButton("tooltip:=Multiple selection","index:=7").Click
Wait(2)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_MI04_1000_ARTICLE,True)'select
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_MI04_3010_TABLECELL_SINGLE_VALUE_1,True)'select
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call SelectCheckbox("IM_SELKB","1",DT_MI04_1000_PHYS_INVENTORY_DOCUMENTS,False)
Call SelectCheckbox("IM_SELKP","1",DT_MI04_1000_PHYS_INVENTORY_ITEMS,False)
Call SelectCheckbox("P_GRID","1","ON",False)

Call SetTextbox("Site","IM_WERKS-LOW","",DT_MI04_1000_SITE,False)
Call SetTextbox("Storage Location","IM_LGORT-LOW","",DT_MI04_1000_STORAGE_LOCATION,False)
Call SetTextbox("Physical Inventory Document","IM_IBLNR-LOW","",DT_MI04_1000_PHYSICAL_INVENTORY_DOCUMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckbox("IX-SELB1","1",DT_MI04_0100_NOT_YET_COUNTED,True)
Call SelectCheckbox("IX-SELB2","1",DT_MI04_0100_PARTIALLY_COUNTED,True)
Call SelectCheckbox("IX-SELB3","1",DT_MI04_0100_ALL_COUNTED,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

Call SelectCheckbox("IX-SELP2","1",DT_MI04_0101_ONLY_CNTD,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Article","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyGridCellContent("",1,"Site","",DT_MI04_0120_CHECK_TEXT_OF_RW22)
Call VerifyGridCellContent("",1,"Storage Location","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyGridCellContent("",1,"Book quantity","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call VerifyGridCellContent("",1,"Qty Counted","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC3)
Call VerifyGridCellContent("",1,"Difference Quantity","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC4)

Call VerifyGridCellContent("",2,"Article","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC5)
Call VerifyGridCellContent("",2,"Site","",DT_MI04_0120_CHECK_TEXT_OF_RW22_OCC1)
Call VerifyGridCellContent("",2,"Storage Location","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC6)
Call VerifyGridCellContent("",2,"Book quantity","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC7)
Call VerifyGridCellContent("",2,"Qty Counted","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC8)
Call VerifyGridCellContent("",2,"Difference Quantity","",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME_OCC9)

Call ClickButtonIfExist("Select All   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Post Difference   \(Ctrl\+Shift\+F9\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Phys\. Inventory Doc\.","RM07I-IBLNR","",DT_MI04_0701_PHYS_INVENTORY_DOC_OCC1,False)
Call SetTextbox("Fiscal Year","RM07I-GJAHR","",DT_MI04_0701_FISCAL_YEAR_OCC1,False)
Call SetTextbox("Posting Date","RM07I-BUDAT","",DT_MI04_0701_POSTING_DATE,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTextStatusBar("DT_MI04_0120_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'verify
Call VerifyStatusBar(DT_MI04_0120_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call ClickButton("Back   \(F3\)",False)
'
''''----------------------Tcode MI24----------------------------
'Enter the Tcode
Call SetTcode(DT_MI04_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MI04_1000_OKCD)
'Capture the screenshot
Call TakeScreenShot()

'multilpe selection for article
SAPGuiSession("transaction:=MI24").SAPGuiWindow("transaction:=MI24").SAPGuiButton("tooltip:=Multiple selection","index:=7").Click
Wait(2)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_MI04_1000_ARTICLE_OCC1,True)'select
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_MI04_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)'select
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call SelectCheckbox("IM_SELKB","1",DT_MI04_1000_PHYS_INVENTORY_DOCUMENTS_OCC1,False)
Call SelectCheckbox("IM_SELKP","1",DT_MI04_1000_PHYS_INVENTORY_ITEMS_OCC1,False)
Call SelectCheckbox("P_GRID","1","ON",False)

Call SetTextbox("Site","IM_WERKS-LOW","",DT_MI04_1000_SITE_OCC1,False)
Call SetTextbox("Storage Location","IM_LGORT-LOW","",DT_MI04_1000_STORAGE_LOCATION_OCC1,False)
Call SetTextbox("Physical Inventory Document","IM_IBLNR-LOW","",DT_MI04_1000_PHYSICAL_INVENTORY_DOCUMENT_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckbox("IX-SELB1","1",DT_MI04_0101_NOT_YET_COUNTED,True)
Call SelectCheckbox("IX-SELB2","1",DT_MI04_0100_PARTIALLY_COUNTED_OCC1,True)
Call SelectCheckbox("IX-SELB3","1",DT_MI04_0100_ALL_COUNTED_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

Call SelectCheckbox("IX-SELP2","1",DT_MI04_0101_ONLY_CNTD_OCC1,True)
Call SelectCheckbox("IX-SELP3","1",DT_MI04_0101_POSTED,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickCellGuiGrid("","","Phys. Inventory Doc.",1,"Article",DT_MI04_0120_CHECK_TEXT_OF_NO_NAME,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Physical Inventory History   \(F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",LCase(DT_CHECK_PHYS_INV_STATUS),False)
Call VerifyTextBoxContent("Article Doc.","ISEG-MBLNR","",DT_CHECK_CREATED_ARTICLE_DOC,False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

