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

''Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : 
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Procurement of consumables from DC to Headquarter_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName=""
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


'''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'---------------------------------me21n-------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()    
Call TakeScreenshot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1105_MEPO_TOPLINEBSART)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_SUPPLYING_SITE,False) 
Call TakeScreenShot

Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)  
Call TakeScreenShot
Call PressEnter() 

Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","A","1","","",DT_ME21N_1211_TABLECELL_A_0,False) 
Call PressEnter() 
wait(2)
Call ClickButtonIfExist("Continue",True)
wait 2
Call ClickButtonIfExist("Continue",True)
wait 2
Call TakeScreenShot

Call SetTextbox("Cost Center","COBL-KOSTL","",DT_ME21N_1101_COST_CENTER,False)  
Call PressEnter() 
Call TakeScreenShot

Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)  
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","A","2","","",DT_ME21N_1211_TABLECELL_A_1,False) 
Call PressEnter() 
wait(2)
Call ClickButtonIfExist("Continue",True)
wait 2
Call ClickButtonIfExist("Continue",True)
wait 2
Call TakeScreenShot

Call SetTextbox("Cost Center","COBL-KOSTL","",DT_ME21N_1101_COST_CENTER_OCC1,False)  
Call PressEnter()
Call TakeScreenShot

Call SelectTab("ITEM_DETAIL","Delivery",False)
Call FocusTextBox("Incoterms","MEPO1313-INCO1",False)
Call ClickButton("Fast Change",False)
Call SetTextbox("Incoterms","MEGUI_MASSCH_ALLOWED_FIELDS-INCO1","",DT_ME21N_0010_INCOTERMS,True)  
Call SelectRadioButton("MEGUI_RANGE-ALL_ITEMS","All items",True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute Changes   \(F8\)",True)
Call TakeScreenShot

Call FocusTextBox("Incoterms","MEPO1313-INCO2",False)
Call ClickButton("Fast Change",False)
Call SetTextbox("Incoterms 2","MEGUI_MASSCH_ALLOWED_FIELDS-INCO2","",DT_ME21N_0010_INCOTERMS_2,True)  
Call SelectRadioButton("MEGUI_RANGE-ALL_ITEMS","All items",True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute Changes   \(F8\)",True)
Call TakeScreenShot

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_DELIVERY_DATE),False)  
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_DELIVERY_DATE),False) 

Call ClickButton("Save   \(Ctrl\+S\)",False)
wait(2)
Call ClickButtonIfExist("Continue",True)
wait 2
Call ClickButtonIfExist("Continue",True)
wait 2
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot

Call GetStatusBar("item2","DT_PO_OUTPUT")
Call VerifyStatusBar("STO Retail created under the number"&DT_PO_NUMBER_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'-----------------------------ME23N-----------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()    
Call TakeScreenshot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER,True)
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order",True)
Call ClickButtonIfExist("Other Document   \(Enter\)",True)

' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Doc\. date","MEPO_TOPLINE-BEDAT","",ConvertDate(DT_DELIVERY_DATE),False)


Call LogOff()
Call FinalStatus ()
