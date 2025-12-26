
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from Third party customers - Integration_p2
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
'.................Test Script Name :Test_Returns from Third party customers - Integration_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Returns from Third party customers - Integration_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Wholesale with WMS delivery - Integration_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VA01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


'Enter the Details
Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,FALSE)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,FALSE)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,FALSE)
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,FALSE)
Call TakeScreenShot()

'Click on Create with Reference
Call ClickButtonIfExist("Create with Reference   \(F8\)",False)
Wait(2)

'Navigate to the Bill Doc  tab
Call SelectTab("MYTABSTRIP","BillDoc",True)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Billing Document","VBRK-VBELN","",DT_VA01_0304_BILLING_DOCUMENT,True)
Call TakeScreenShot()

'Click on Copy
Call ClickButtonIfExist("Copy   \(F5\)",True)
Wait(2)

Call SetComboByKey("GS_MSR_SALES_ITEM-REFUND_CONTROL",DT_VA01_4448_REFUND_CONTROL)

Call SetComboByKey("GS_MSR_SALES_ITEM-REFUND_TYPE"," ")

Call SelectRowGuiTableByRow("SAPMV45ATCTRL_U_ERF_AUFTRAG",1,False)


Call SelectMenuBar("Edit;Fast change of...;Plant...")

Call SetTextbox("Site","RV45A-S_WERKS","",DT_VA01_0257_SITE,True)
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F7\)",True)
Wait(2)


'Navigate to the Returns  tab
Call SelectTab("TAXI_TABSTRIP_OVERVIEW","Returns",False)
Wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("Display header details",False)
Wait(2)

'Navigate to the Conditions  tab
Call SelectTab("TAXI_TABSTRIP_HEAD","Conditions",False)
Wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("Update Prices",False)
Wait(2)

Call SetFocusGuiLabel("G",11,104,True)

Call ClickButtonIfExist("Copy   \(Enter\)",True)
Wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Save the Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

'Verify the Status Bar msg
Call GetStatusBar("item2","DT_ADV_RETURN_ORDER")
Call VerifyStatusBar("AB: Advanced Returns "&DT_ADV_RETURN_ORDER&" has been saved")


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


