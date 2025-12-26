	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_02IMP00_007_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



'
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

gstrTestCaseName = "Test_PRE_02IMP00_007_P1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_02IMP00_006_P1.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

'//-----------------------------------MIGO -----------------------------------

call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)  
Call TakeScreenshot()
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call PressEnter() 
Call TakeScreenShot


Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
wait(3)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False)   
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)    
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)   
Call TakeScreenShot
Call PressEnter()     ' - Line (26)

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)  
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_ME21N_1211_TABLECELL_OUN_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",REPLACE(DT_ME21N_1211_TABLECELL_DELIV_DATE_0,"/","."),False)
Call TakeScreenShot

Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False)    
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)   
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)   
Call SetTableData("SAPLMEGUITC_1211","OUn","2","","",DT_ME21N_1211_TABLECELL_OUN_1,False)   
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",REPLACE(DT_ME21N_1211_TABLECELL_DELIV_DATE_1,"/","."),False)    
Call TakeScreenShot

Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False)  
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)   
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False)   
Call SetTableData("SAPLMEGUITC_1211","OUn","3","","",DT_ME21N_1211_TABLECELL_OUN_2,False)     
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",REPLACE(DT_ME21N_1211_TABLECELL_DELIV_DATE_2,"/","."),False)   
Call TakeScreenShot

Call SetTableData("SAPLMEGUITC_1211","Article","4","","",DT_ME21N_1211_TABLECELL_ARTICLE_3,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","4","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_3,False)   
Call SetTableData("SAPLMEGUITC_1211","Plnt","4","","",DT_ME21N_1211_TABLECELL_SITE_3,False)   
Call SetTableData("SAPLMEGUITC_1211","OUn","4","","",DT_ME21N_1211_TABLECELL_OUN_3,False)    
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","4","","",REPLACE(DT_ME21N_1211_TABLECELL_DELIV_DATE_3,"/","."),False)  
Call TakeScreenShot

Call SetTableData("SAPLMEGUITC_1211","Article","5","","",DT_ME21N_1211_TABLECELL_ARTICLE_4,False)  
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","5","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_4,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","5","","",DT_ME21N_1211_TABLECELL_SITE_4,False)  
Call SetTableData("SAPLMEGUITC_1211","OUn","5","","",DT_ME21N_1211_TABLECELL_OUN_4,False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","5","","",REPLACE(DT_ME21N_1211_TABLECELL_DELIV_DATE_4,"/","."),False)   
Call TakeScreenShot

Call PressEnter()
Call TakeScreenShot

Call SelectTab("ITEM_DETAIL","Delivery",False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("ITEM_DETAIL","Delivery",False)
Call SetTextbox("Latest GR Date","MEPO1313-LEWED","",replace(DT_ME21N_1313_LATEST_GR_DATE,"/","."),False)  
Call TakeScreenShot
Call PressEnter()

call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(3)
Call TakeScreenShot
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call SelectTab("HEADER_DETAIL","Conditions",False)
Call TakeScreenShot
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)

call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call ClickButton("Insert Row",False)
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)

Call GetNumberOfRows("SAPLV69ATCTRL_KONDITIONEN", "", "NUMBER_OF_ROWS")

Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","CnTy",NUMBER_OF_ROWS-1,"","",DT_ME21N_6201_TABLECELL_CNTY_1,False)
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN"," Amount",NUMBER_OF_ROWS-1,"","",DT_ME21N_6201_TABLECELL_AMOUNT_1,False)
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","CnTy",NUMBER_OF_ROWS-2,"","",DT_ME21N_6201_TABLECELL_CNTY_2,False)
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN"," Amount",NUMBER_OF_ROWS-2,"","",DT_ME21N_6201_TABLECELL_AMOUNT_2,False)
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
wait 10

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Wait 10
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)

Call SelectCellGuiTable("SAPLV69ATCTRL_KONDITIONEN", "CnTy", "CnTy", "ZHC5", False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call FindRowNumber("SAPLV69ATCTRL_KONDITIONEN","CnTy","ZHC5","DT_ROW_NUMBER_OUTPUT")
wait 5
Call TakeScreenShot
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait 5
Call FindRowNumber("SAPLV69ATCTRL_KONDITIONEN","CnTy","ZHC5","DT_ROW_NUMBER_OUTPUT")
Call TakeScreenShot
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectRowGuiTableByRow("SAPLV69ATCTRL_KONDITIONEN",DT_ROW_NUMBER_OUTPUT+1,False)
wait 5
Call TakeScreenShot

CAll ClickButton("Condition detail",False)
Call SetTextbox("Supplier","KOMV-LIFNR","",DT_ME21N_0640_VENDOR,False) 
Call PressEnter()
Call TakeScreenShot
call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot


Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Wait 10
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)

'Call SelectCellGuiTable("SAPLV69ATCTRL_KONDITIONEN", "CnTy", "CnTy", "ZHC8", False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot
Call FindRowNumber("SAPLV69ATCTRL_KONDITIONEN","CnTy","ZHC8","DT_ROW_NUMBER_OUTPUT")
wait 5
Call TakeScreenShot
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait 5
Call TakeScreenShot
call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectRowGuiTableByRow("SAPLV69ATCTRL_KONDITIONEN",DT_ROW_NUMBER_OUTPUT+1,False)
wait 5
Call TakeScreenShot

CAll ClickButton("Condition detail",False)
Call SetTextbox("Supplier","KOMV-LIFNR","",DT_ME21N_0640_VENDOR_OCC1,False) 
Call PressEnter()
Call TakeScreenShot
call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False) 

Call TakeScreenShot
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot

Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item2","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number "&DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()
