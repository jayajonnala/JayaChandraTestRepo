
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06ARM04_001_Goods_Receipt_for_Return_Order_Franchisee_Custom_Var2
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
'.................Test Script Name :Test_06ARM04_001_Goods_Receipt_for_Return_Order_Franchisee_Custom_Var2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_06ARM04_001_Franchisee_Custom_Var2"
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

''----------------------Tcode MSR_TRC_C----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Sales Document","SO_VBELN-LOW","",DT_MSR_TRC_C_0302_SALES_DOCUMENT,FALSE)
Call TakeScreenShot()


'Click on Execution
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'''Call VerifyGridCellContent("",1,"Processing Status",0,DT_MSR_TRC_C_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROC_STATUS_ICON)
'''Call VerifyGridCellContent("",1,"Processing Status",1,DT_MSR_TRC_C_0300_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROC_STATUS_ICON_OCC1)

Call VerifyGridCellContent("",1,"Processing Status",0,DT_MSR_TRC_C_0300_CHECK_PROC_STATUS_ICON)
Call VerifyGridCellContent("",1,"Processing Status",1,DT_MSR_TRC_C_0300_CHECK_PROC_STATUS_ICON_OCC1)

'Call ActivateNodeGuiTree(0,"Customer return with shipment to site;Returns Delivery")
Call ActivateItemGuiTree(0,"#1;Returns Delivery","Returns Delivery")

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

'Click on Outbound Delivery;Display <-> Change
Call SelectMenuBar("Outbound Delivery;Display <-> Change")

Call GetTextboxValue("LIKP-VBELN","0","DT_OUTBOUND_DELV_NO_OUTPUT",False)

'Save the Post Goods Receipt
Call ClickButton("Post Goods Receipt   \(Shift\+F8\)",False) 
Wait(2)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)


'Verify the Status Bar message
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MSR_TRC_C_0300_CHECK_TEXT_OF_STATUSBAR)

'Click on Execution
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")

'Call ActivateNodeGuiTree(0,"Customer return with shipment to site;Goods Receipt")

Call ActivateItemGuiTree(0,"#1;Goods Receipt","Goods Receipt")

Call SelectTab("TS_GOITEM","Where",False)
Wait(1)
Call TakeScreenShot()
Call VerifyTextBoxContent("Storage Location","GOITEM-LGORT",0,DT_MSR_TRC_C_0325_CHECK_TEXT_OF_STORAGE_LOCATION,False)

'Navigate to Quantity Tab
Call SelectTab("TS_GOITEM","Quantity",False)
Wait(1)
Call TakeScreenShot()

Call VerifyTextBoxContent("Qty in Unit of Entry","GOITEM-ERFMG",0,DT_MSR_TRC_C_0315_CHECK_TEXT_OF_QTY_IN_UNIT_OF_ENTRY,False)
Call VerifyTextBoxContent("Qty in SKU","GOITEM-MENGE",0,DT_MSR_TRC_C_0315_CHECK_TEXT_OF_QTY_IN_SKU,False)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)
Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


