
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0281_Industry field same art from 2 vend one with 9percent VAT and one with 0percent VAT_P1
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

''gstrTestCaseName = "Test_P2P_01_01_0281_Industry field same art from 2 vend one with 9percent VAT and one with 0percent VAT_P1"
gstrTestCaseName = "Test_P2P_01_01_0281_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''-''''----------Set TCode-ME21N----------------'''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1105_MEPO_TOPLINEBSART)

Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot

'''''---------VENDOR 1------------'''''''''

Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)         
Call PressEnter()    
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)     
Call FocusTextBox("Company Code","MEPO1222-BUKRS",False)     
Call PressEnter()     
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)     

'NewDate= CSTR(Day(DT_DEL_DATE)) +"."+ Cstr (Month(DT_DEL_DATE)) + "."+ Cstr(Year(DT_DEL_DATE))
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_DEL_DATE),False)     

Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)     
Call PressEnter()     

Call TakeScreenShot

Call SelectTab("ITEM_DETAIL","Invoice",False)
Call SetTextbox("Tax Code","MEPO1317-MWSKZ","",DT_ME21N_1317_TAX_CODE,False)
'Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",false)
Call ClickButtonIfExist("Save",True)

Call TakeScreenShot

Call GetStatusBar("item2","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number "&DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)

Call ClickButton("Create   \(F6\)",false)
Call PressEnter() 
Call SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1105_MEPO_TOPLINEBSART_OCC1)

'Call SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1211_TABLECELL_ARTICLE_0_OCC1)

Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR_OCC1,False)          
Call PressEnter()    
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG_OCC1,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP_OCC1,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE_OCC1,False)     
Call FocusTextBox("Company Code","MEPO1222-BUKRS",False)     
Call PressEnter()     

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0_OCC1,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0_OCC1,False)     

NewDate= CSTR(Day(DT_DEL_DATE)) +"."+ Cstr (Month(DT_DEL_DATE)) + "."+ Cstr(Year(DT_DEL_DATE))
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",NewDate,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0_OCC1,False)    
Call TakeScreenShot

Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)
Call SelectTab("ITEM_DETAIL","Invoice",False)
Call SetTextbox("Tax Code","MEPO1317-MWSKZ","",DT_ME21N_1317_TAX_CODE_OCC1,False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",false)
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot

Call GetStatusBar("item2","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number "&DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT)
'
'
'''''''----------Set TCode-MIGO----------------'''''
'
Call SetTcode(DT_ME21N_0014_OKCD)     
Call PressEnter() 
Call TakeScreenShot
Call CheckTCodeScreen(DT_ME21N_0014_OKCD)

Call SetComboByKey("GODYNPRO-ACTION", DT_ME21N_1105_MEPO_ACTION)
Call PressEnter() 
Call SetComboByKey("GODYNPRO-REFDOC", DT_ME21N_1105_MEPO_ACTION_REFDOC)

Call PressEnter()

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call PressEnter()

Call ClickButtonIfExist("Open detail data",False)
CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call PressEnter() 

Call ClickButton("Check Entries   \(F7\)",false)
Call VerifyStatusBar(DT_ME21N_0001_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Post Document   \(Shift\+F11\)",false)

Call GetStatusBar("item1","DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")


Call ClickButton("Restart   \(F5\)",false)
Call PressEnter() 

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT, False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE_OCC1,False)
Call PressEnter()

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call PressEnter() 

Call ClickButton("Check Entries   \(F7\)",false)
Call VerifyStatusBar(DT_ME21N_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("item1","DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Article document "&DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" posted")

Call LogOff()

Call FinalStatus ()



