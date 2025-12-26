

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_5_3_1 Article inspection for Return Order_P3
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

gstrTestCaseName = "Test_5_3_1 Article Order_P3"
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


'--------------------------------VA01-----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)   
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False) 
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False) 
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False) 
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False) 
Call PressEnter() 
Call TakeScreenShot()
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article","1","","",DT_VA01_4900_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity","1","","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call PressEnter() 
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call GetStatusBar("item2","DT_RETURN_SO_NUMBER_OUTPUT")
Call VerifyStatusBar("AB: Advanced Returns "& DT_RETURN_SO_NUMBER_OUTPUT &" has been saved")

''--------------------------------va03-----------------------------
Call SetTcode(DT_VA01_0100_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_VA01_0100_OKCD)

Call SetTextbox("Order","VBAK-VBELN","",DT_RETURN_SO_NUMBER_OUTPUT,False) 
Call TakeScreenShot()
Call ClickButton("Search",False)
Call TakeScreenShot()
Call ClickButton("Go to Returns Overview",False)

Call ClickButtonIfExist("Refresh   \(F8\)",False)
Call ActivateItemGuiTree(0,"#1;Returns Delivery","Returns Delivery")
Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)  
Call GetTextboxValue("LIKP-VBELN","","DT_VA01_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT",False)
Call ClickButton("Post Goods Receipt   \(Shift\+F8\)",False) 
Call VerifyStatusBar("Adv.Returns Delivery "& DT_VA01_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT &" has been saved")
Call ClickButton("Refresh   \(F8\)",False)
Call ClickButton("Refresh   \(F8\)",False)



''--------------------------------msr_inspwh-----------------------------
Call SetTcode(DT_VA01_0100_OKCD_OCC1) 
Call PressEnter()     
Call CheckTCodeScreen(DT_VA01_0100_OKCD_OCC1)

Call SetTextbox("Delivery","INBD2-LOW","",DT_VA01_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True) 

Call SelectTab("TABSTRIP_INSP","Item",False)
Call SetTextbox("Inspection Code","MSR_S_INSP_UI_ITM-DEC_CODE","",DT_VA01_2010_INSPECTION_CODE,False)
Call SetComboByKey("MSR_S_INSP_UI_ITM-FU_CODE",DT_VA01_2010_FOLLOWUP_ACTIVITY)
Call PressEnter()
Call SelectCheckbox("MSR_S_INSP_UI_ITM-REFUND_RELEASED",0,DT_VA01_2010_RELEASE_CREDIT_MEMO_REQUEST,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Save and Confim LFU   \(Ctrl\+F8\)",False)
Call VerifyStatusBar("Inspection for delivery "& DT_VA01_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT &" saved")
Call ClickButtonIfExist("Open Returns Overview   \(Ctrl\+F12\)",False)
Wait 15
Call ClickButton("Refresh   \(F8\)",False)
Wait 15
Call ClickButton("Refresh   \(F8\)",False)
Wait 15
Call ClickButton("Refresh   \(F8\)",False)
Call ActivateItemGuiTree(0,"#1;Credit Memo Request","Credit Memo Request")
Call SelectMenuBar("Sales document;Billing")
Call TakeScreenShot()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_CM_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "& DT_CM_NUMBER_OUTPUT &" has been saved")

''--------------------------------va03-----------------------------
Call SetTcode(DT_VA01_0102_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_VA01_0102_OKCD)

Call SetTextbox("Order","VBAK-VBELN","",DT_RETURN_SO_NUMBER_OUTPUT,False)
Call TakeScreenShot()
Call ClickButton("Search",False) 
Call TakeScreenShot()
Call ClickButton("Go to Returns Overview",False)
Call ClickButton("Refresh   \(F8\)",False)
Wait 2
Call VerifyTextBoxContent("Processing Status","MSR_S_SCR_ITEM_SD-PROC_STATUS_ICON","",LCASE(DT_VA01_0110_CHECK_TEXT_OF_PROCESSING_STATUS),False)


Call LogOff()
Call FinalStatus ()



