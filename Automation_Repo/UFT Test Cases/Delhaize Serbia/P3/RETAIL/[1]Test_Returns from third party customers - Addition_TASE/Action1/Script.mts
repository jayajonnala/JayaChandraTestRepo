
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from third party customers - Addition
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
'.................Test Script Name :Test_Returns from third party customers - Addition
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Returns from third party customers - Addition"
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
Call SelectTab("MYTABSTRIP","Order",True)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("Order","LV45C-VBELN","",DT_VA01_0301_ORDER,True)
Call TakeScreenShot()

'Click on Copy
Call ClickButtonIfExist("Copy   \(F5\)",True)
Wait(2)

'Save the Details
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)

'Verify the Status Bar msg
Call GetStatusBar("item2","DT_REQ_CREDIT_MEMO_OUTPUT")
Call VerifyStatusBar("AB: ARM Credit Req. "&DT_REQ_CREDIT_MEMO_OUTPUT&" has been saved")

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

''----------------------Tcode VF01----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_VA01_0101_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VA01_0101_OKCD)

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document","1","","",DT_VA01_0102_TABLECELL_DOCUMENT_0,False)
Call TakeScreenShot()
Call PressEnter()
Wait(2)


'Click on Save
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)


'Validate If Document No is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" has been saved")

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'----------------------Tcode VF03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_VA01_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_VA01_0100_OKCD)

'Display the Invoice Details
Call SetTextbox("Billing document","VBRK-VBELN","",DT_VA01_0101_BILLING_DOCUMENT,False)
Call TakeScreenShot()

'Navigate to Billing document;Issue Output To
Call SelectMenuBar("Billing document;Issue Output To")
Call TakeScreenShot()

'Select Row
Call SelectRowGuiTable("SAPLVMSGTABCONTROL","Message type","ZGRS",True)
Call TakeScreenShot()

'Click on Print
Call ClickButton("Print preview   \(Ctrl\+Shift\+F1\)",True)
Wait(5)
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Wait(2)

'Click on Accounting
Call ClickButtonIfExist("Accounting overview   \(F6\)",False)
Wait(2)
Call TakeScreenShot()

'Call DoubleClickGuiGridCell("Documents in Accounting",0,1,"Document Number",True)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",1,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call TakeScreenShot()

Call  GetTextboxValue("BKPF-BELNR",0,"DT_ACCOUNTING_DOC_NO",False)

Call SelectColumnGuiGrid("",0,"Account",False)
Call ClickButtonToolBar("&MB_FILTER",0)

Call SetTextboxPopupIfExist("%%DYN001-LOW","Account",DT_VA01_1105_ACCOUNT)
Wait(1)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
Wait(2)


Call VerifyGridCellContent("",1,"Profit Center",0,DT_VA01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


