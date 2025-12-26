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

  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Reload DataSheet to updates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_DSD - Rprod - with PO CC RS01_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 13th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_DSD - Rprod - with PO CC RS01_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_DSD - Rprod - with PO CC RS01_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
'Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode ME21N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)

Call SelectTab("HEADER_DETAIL","Org. Data",False)

''Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
wait(1)
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(1)
Call SelectTab("HEADER_DETAIL",DT_ME21N_1102_ORG_DATA,False)
wait(1)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False) 

Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
wait(1)
'Call SetTableData("SAPLMEGUITC_1211","Returns Item","1","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_0,False)
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
'Capture the screenshot
Call TakeScreenShot()

'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
'Call SetTextbox("Doc\. date","MEPO_TOPLINE-BEDAT","",Replace((DT_ME21N_1105_DOC_DATE),"/","."),False) 
'Capture the screenshot
Call TakeScreenShot()

'Click on Save Buton
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(1)
Call ClickButtonIfExist("Continue",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Save",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
'Call GetStatusBar("item1","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_ME21N_PO_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)

Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
''''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_ME21N_0100_OKCD) 
Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

'Enter the PO Number and Press Enter
'Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
'Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetComboByKey("GODYNPRO-ACTION",DT_ME21N_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_ME21N_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",Replace((DT_ME21N_0110_DOCUMENT_DATE),"/","."),False) 
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",Replace((DT_ME21N_0110_POSTING_DATE),"/","."),False) 
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_ME21N_0200_CHECK_SELECTED_OF_TABLECELL_OK_0,False)
Call VerifyCheckBoxValue("GODYNPRO-DETAIL_TAKE",DT_ME21N_0200_CHECK_SELECTED_OF_TABLECELL_OK_0)
'Capture the screenshot
Call TakeScreenShot()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_ME21N_DOCUMENT_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_ME21N_0001_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


