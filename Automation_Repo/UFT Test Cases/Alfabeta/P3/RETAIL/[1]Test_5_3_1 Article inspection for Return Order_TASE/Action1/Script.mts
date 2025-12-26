

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_5_3_1 Article inspection for Return Order
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_5_3_1 Article Return Order"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''


''--------------------------------ME21N-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Call SetTextbox("Supplying Site","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_SUPPLYING_SITE,False)   
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False) 
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()

Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True) 
wait 2
Call GetStatusBar("item2","DT_PO_NUM_OUTPUT")
Call VerifyStatusBar("STO Return Order created under the number " & DT_PO_NUM_OUTPUT )


call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_PO_NUM_OUTPUT,True)    
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
Call ClickButton("Other Document   \(Enter\)",True)  

Call SelectTab("ITEM_DETAIL","Retail",False)
Call TakeScreenShot()

Call SelectTab("ITEM_DETAIL","Returns",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Open Returns Overview",False) 
Call TakeScreenShot()

Call ActivateItemGuiTree(0,"#1;Outbound Delivery","Outbound Delivery")
Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)  
Call GetTextboxValue("LIKP-VBELN","","DT_ME21N_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT",False)
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False) 
Call VerifyStatusBarExist("Adv.Returns Delivery "& DT_ME21N_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT &" has been saved")
Call ClickButton("Refresh   \(F8\)",False)
wait 15
Call ClickButton("Refresh   \(F8\)",False)
wait 15
Call ClickButton("Refresh   \(F8\)",False)

Call ActivateItemGuiTree(0,"#1;Inbound Delivery","Inbound Delivery")
Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)  
Call GetTextboxValue("LIKP-VBELN","","DT_ME21N_1602_CHECK_TEXT_OF_INBOUND_DELIV_OUTPUT",False)
Call ClickButton("Post Goods Receipt   \(Shift\+F8\)",False) 
 ' VerifyStatusBarExist(Content)
Call VerifyStatusBarExist("Goods receipt for inbound delivery "& DT_ME21N_1602_CHECK_TEXT_OF_INBOUND_DELIV_OUTPUT &" posted")
Call ClickButton("Refresh   \(F8\)",False)

''--------------------------------MSR_INSPWH-----------------------------
Call SetTcode(DT_ME21N_0100_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_ME21N_0100_OKCD)

Call SetTextbox("Delivery","INBD2-LOW","",DT_ME21N_1602_CHECK_TEXT_OF_INBOUND_DELIV_OUTPUT,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True) 

Call SelectTab("TABSTRIP_INSP","Item",False)
Call SetTextbox("Inspection Code","MSR_S_INSP_UI_ITM-DEC_CODE","",DT_ME21N_2010_INSPECTION_CODE,False)
Call SetComboByKey("MSR_S_INSP_UI_ITM-FU_CODE",DT_ME21N_2010_FOLLOWUP_ACTIVITY)
Call PressEnter()
Call TakeScreenShot()

Call ClickButtonIfExist("Save and Confim LFU   \(Ctrl\+F8\)",False)
Call VerifyStatusBar("Inspection for delivery "& DT_ME21N_1602_CHECK_TEXT_OF_INBOUND_DELIV_OUTPUT &" saved")
Call ClickButtonIfExist("Open Returns Overview   \(Ctrl\+F12\)",False)
Call ClickButton("Refresh   \(F8\)",False)
wait 15
Call ClickButton("Refresh   \(F8\)",False)
wait 15
Call ClickButton("Refresh   \(F8\)",False)
Call VerifyTextBoxContent("Processing Status","MSR_S_SCR_HEAD_DLV-PROC_STATUS_ICON","",LCASE(DT_ME21N_0113_CHECK_TEXT_OF_PROCESSING_STATUS),False)

Call LogOff()
Call FinalStatus ()
