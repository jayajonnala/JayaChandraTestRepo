
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Import - regular goods_p1
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Import - regular goods_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Import - regular goods_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Import - regular goods_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'----------------------Tcode ME21N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Wait(1)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Wait(1)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

'Enter Order Details'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Navigate to Delivery Tab and Enter Latest GR Date
Call SelectTab("ITEM_DETAIL","Delivery",False)
Wait(1)
Call SetTextbox("Latest GR Date","MEPO1313-LEWED","",DT_GR_DATE,FALSE)
Call TakeScreenShot()


'Navigate to Invoice Tab and Select the GR-Bsd IV checkbox
Call SelectTab("ITEM_DETAIL","Invoice",False)
Wait(1)
Call VerifyCheckBoxValue("MEPO1317-WEBRE","OFF")
Call TakeScreenShot()
'
''Navigate to Condition Tab
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Conditions",False)
Wait(1)

'Click on Inser Roz
Call ClickButton("Insert Row",False) 
Wait(2)
'
''Enter the Details
Call SetTableDataNoRef("SAPLV69ATCTRL_KONDITIONEN","CnTy",31,DT_ME21N_6201_TABLECELL_CNTY_1,False)
Call SetTableDataNoRef("SAPLV69ATCTRL_KONDITIONEN","Amount",31,DT_ME21N_6201_TABLECELL_AMOUNT_1,False)
Call SetTableDataNoRef("SAPLV69ATCTRL_KONDITIONEN","Crcy",31,DT_ME21N_6201_TABLECELL_CRCY_1,False)
Call PressEnter()
Wait(1)
Call TakeScreenShot()


Call SelectRowGuiTable("SAPLV69ATCTRL_KONDITIONEN","CnTy",DT_ME21N_6201_TABLECELL_CNTY_1,False)
'Click on Condition details
Call ClickButton("Condition detail",False) 
Wait(2)


'Call SetTextbox("Vendor","KOMV-LIFNR","",DT_ME21N_0640_VENDOR,FALSE)
Call SetTextboxNoLabel("KOMV-LIFNR","",DT_ME21N_0640_VENDOR,FALSE)
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Wait(2)
Call ClickButtonIfExist("Save",True)

'Validate If STO Retail Order is generated
Call GetStatusBar("item2","DT_PO_RETAIL_ORDER_OUTPUT")
VerifyStatusBar("Standard PO Retail created under the number " & DT_PO_RETAIL_ORDER_OUTPUT)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()


'*********************************************End Of Script*********************************************************************

