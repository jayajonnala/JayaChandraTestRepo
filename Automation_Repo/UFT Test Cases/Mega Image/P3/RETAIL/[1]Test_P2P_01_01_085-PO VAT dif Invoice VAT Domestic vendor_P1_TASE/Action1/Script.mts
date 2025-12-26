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
'.................Test Script Name : Test_P2P_01_01_085-PO VAT dif Invoice VAT Domestic vendor_P1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 2nd June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_085 vendor_P1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_085-PO VAT  dif Invoice VAT  Domestic vendor_P1_TASE.xls"
'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode ME21N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

''Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
wait(1)
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(1)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
wait(1)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)   
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False) 

Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
wait(1)
'Call SetTableData("SAPLMEGUITC_1211","Returns Item","1","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_0,False)
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False)
Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False)
Call SetTableData("SAPLMEGUITC_1211","Article","4","","",DT_ME21N_1211_TABLECELL_ARTICLE_3,False)

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","4","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_3,False)

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(Date),False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(Date),False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(Date),False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","4","","",ConvertDate(Date),False)

Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","4","","",DT_ME21N_1211_TABLECELL_SITE_3,False)
'Capture the screenshot
Call TakeScreenShot()

For Iterator = 1 To 9 Step 1
	Call PressEnter()  
Next  

'Capture the screenshot
Call TakeScreenShot()
Call SelectTab("ITEM_DETAIL",DT_ME21N_1303_INVOICE,False)

'get data
Call SetComboByKey("DYN_6000-LIST","   2")
Wait(3)
Call GetTextboxValue("MEPO1317-MWSKZ","","DT_ME21N_1317_GET_TEXT_OF_TAX_CODE_OCC2_OUTPUT",False)

Call ClickButtonIfExist("Previous item",False)
Wait(3)
Call GetTextboxValue("MEPO1317-MWSKZ","","DT_ME21N_1317_GET_TEXT_OF_TAX_CODE_OUTPUT",False)

Call SetComboByKey("DYN_6000-LIST","   3")
Wait(3)
Call GetTextboxValue("MEPO1317-MWSKZ","","DT_ME21N_1317_GET_TEXT_OF_TAX_CODE_OCC4_OUTPUT",False)

Call SetComboByKey("DYN_6000-LIST","   4")
Wait(3)
Call GetTextboxValue("MEPO1317-MWSKZ","","DT_ME21N_1317_GET_TEXT_OF_TAX_CODE_OCC6_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'verify
Call SetComboByKey("DYN_6000-LIST","   2")
Wait(3)
Call VerifyTextBoxContent("Tax Code","MEPO1317-MWSKZ","",DT_ME21N_1317_CHECK_TEXT_OF_TAX_CODE_OCC3,False)

Call ClickButtonIfExist("Previous item",False)
Wait(3)
Call VerifyTextBoxContent("Tax Code","MEPO1317-MWSKZ","",DT_ME21N_1317_CHECK_TEXT_OF_TAX_CODE_OCC1,False)

Call SetComboByKey("DYN_6000-LIST","   3")
Wait(3)
Call VerifyTextBoxContent("Tax Code","MEPO1317-MWSKZ","",DT_ME21N_1317_CHECK_TEXT_OF_TAX_CODE_OCC5,False)

Call SetComboByKey("DYN_6000-LIST","   4")
Wait(3)
Call VerifyTextBoxContent("Tax Code","MEPO1317-MWSKZ","",DT_ME21N_1317_CHECK_TEXT_OF_TAX_CODE_OCC7,False)


Call ClickButtonIfExist("Display Messages   \(Shift\+F6\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 

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
Call GetStatusBar("item2","DT_DOCUMENT_NUMBER_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOCUMENT_NUMBER_OUTPUT",DT_DOCUMENT_NUMBER)
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
'Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


