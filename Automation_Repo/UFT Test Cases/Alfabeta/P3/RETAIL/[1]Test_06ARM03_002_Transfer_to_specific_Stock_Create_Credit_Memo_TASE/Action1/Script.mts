
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06ARM03_002_Transfer_to_specific_Stock
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
'.................Test Script Name :Test_06ARM03_002_Transfer_to_specific_Stock
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_06ARM03_002_Transfer_to_specific_Stock_Create_Credit_Memo_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Direct Re- Export without border crossing_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode MSR_INSPWH----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()
'
Call SetTextbox("Delivery","INBD2-LOW","",DT_MSR_INSPWH_0110_DELIVERY,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True) 
Wait(2)
Call TakeScreenShot()

'Navigate to Item Tab
Call SelectTab("TABSTRIP_INSP","Item",False)
Wait(1)
Call TakeScreenShot()

'''Enter the details
Call VerifyTextBoxContent("Article","MSR_S_INSP_UI_ITM-MATNR",0,DT_MSR_INSPWH_2010_CHECK_TEXT_OF_ARTICLE,False)
Call SetTextbox("Inspection Code","MSR_S_INSP_UI_ITM-DEC_CODE","",DT_MSR_INSPWH_2010_INSPECTION_CODE,False)
Call SetComboByKey("MSR_S_INSP_UI_ITM-FU_CODE",DT_MSR_INSPWH_2010_FOLLOWUP_ACTIVITY_OCC2)
Call TakeScreenShot()
Call PressEnter()
'''Click on Save and Confim
Call TakeScreenShot()
Call ClickButton("Save and Confim LFU   \(Ctrl\+F8\)",False) 
Call GetStatusBar("item1","DT_INSPECTION_OUTPUT")
Call VerifyStatusBar("Inspection for delivery "&DT_INSPECTION_OUTPUT&" saved")
'''Click on Open Returns Overview
Call ClickButton("Open Returns Overview   \(Ctrl\+F12\)",False) 
Call TakeScreenShot()
'Updated Credit memo request to Returns Delivery 
Call ActivateNodeGuiTree(0,"Customer return with shipment to plant;Returns Delivery")
Call TakeScreenShot()
'Updated below code to verify T-Code screens as per TAO log
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False) 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()
''----------------------Tcode  MSR_CRD----------------------------
Call SetTcode(DT_MSR_INSPWH_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Delivery","S_VBLVL-LOW","",DT_MSR_INSPWH_0111_DELIVERY,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True) 
Call TakeScreenShot()
Wait(2)
Call TakeScreenShot()
Call ClickButtonToolBar("301_CREATE",0)
'Select Item overview
Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Item overview",False)
'Set Focust Article 2
Call FocusTableCell("SAPMV45ATCTRL_U_ERF_GUTLAST", "Article", 1, "Item", 10, "MSR_S_INSP_UI_ITM-MATNR", False)
'seect Sales/Documenting
Call SelectMenuBar("Sales document;Billing")
'Verify Tcode VF01 screen
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
Call GetStatusBar("item2","DT_ARM_CRED_REQ_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ARM_CRED_REQ_OUTPUT",DT_ARM_CRED_REQ)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MSR_INSPWH_0104_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Wait(2)
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MSR_INSPWH_0102_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Back   \(F3\)",False) 

'''----------------------Tcode MSR_TRC_C----------------------------

Call SetTcode(DT_MSR_INSPWH_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MSR_INSPWH_0100_OKCD_OCC1)
Call TakeScreenShot()
'Updated below as per TAO log
Call SetTextbox("Sales Document","SO_VBELN-LOW","",DT_MSR_INSPWH_0302_SALES_DOCUMENT,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
'Updated below as per TAO log
Call ActivateNodeGuiTree(0,"Customer return with shipment to plant;Credit Memo")
Call GetTextboxValue("VBRK-VBELN",0,"DT_CREDIT_MEMO_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_CREDIT_MEMO_OUTPUT",DT_CREDIT_MEMO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)
Call ClickButton("Back   \(F3\)",False) 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)
Call TakeScreenShot()
Call VerifyGridCellContent("",1,"Logistical Follow-Up Status",0,DT_MSR_INSPWH_CHECK_LOG_STATUS_TEXT)
Call VerifyGridCellContent("",1,"Refunding Status",1,DT_MSR_INSPWH_CHECK_FIN_STATUS_TEXT)
Call TakeScreenShot()


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


