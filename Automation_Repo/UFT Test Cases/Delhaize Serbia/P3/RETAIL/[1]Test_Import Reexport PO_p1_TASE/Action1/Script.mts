'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Import Reexport PO_p1
'.................Author : TCS        
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

gstrTestCaseName = "Test_Import Reexport PO_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-ME21N----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Call SetTextbox("Doc\. date","MEPO_TOPLINE-BEDAT","",ConvertDate(DT_DOC_DATE),False)     
Call TakeScreenShot
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)

Call TakeScreenShot
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)  
Call TakeScreenShot
Call ClickBUttonIfExist("Collapse Items Ctrl\+F6",False)
Call TakeScreenShot
Call SelectTab("ITEM_DETAIL","Texts",False)
Call TakeScreenShot
Call SelectTab("ITEM_DETAIL","Invoice",False)
Call TakeScreenShot
Call SelectCheckBox("MEPO1317-WEBRE",0,DT_ME21N_1317_GRBSD_IV,False)
Call TakeScreenShot
Call SelectTab("ITEM_DETAIL","Delivery",False)
Call TakeScreenShot
Call SetTextbox("Latest GR Date","MEPO1313-LEWED","","31.12.9999",False)
Call ClickButtonIfExist("DYN_4000-BUTTON",False)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call TakeScreenShot
Call SetTableData("SAPLMEGUITC_1211","Net Price","1","","","2,81",False)
Call SelectTab("HEADER_DETAIL","Conditions",False)
Call TakeScreenShot

Call FindRowNumber("SAPLV69ATCTRL_KONDITIONEN", "Description", "", "DT_ROW_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ROW_NUMBER_OUTPUT",DT_ROW_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","CnTy",DT_ROW_NUMBER,"","",DT_ME21N_6201_TABLECELL_CNTY_4,False) 
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","Amount",DT_ROW_NUMBER,"","",DT_ME21N_6201_TABLECELL_AMOUNT_4,False)
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","Crcy",DT_ROW_NUMBER,"","",DT_ME21N_6201_TABLECELL_CRCY_4,False)

Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot() 

Call SelectRowGuiTable("SAPLV69ATCTRL_KONDITIONEN", "CnTy", DT_ME21N_6201_TABLECELL_CNTY_4, False)

Call ClickButton("BT_KODE",False)
Call TakeScreenShot()
Call SetTextboxNoLabel("KOMV-LIFNR","",DT_ME21N_1105_VENDOR,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call GetTextStatusBar("DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call LogOff
Call FinalStatus()
