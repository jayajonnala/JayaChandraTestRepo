
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from Third party customers - Integration_p3
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
'.................Test Script Name :Test_Returns from Third party customers - Integration_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Returns from Third party customers - Integration_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Wholesale with WMS delivery - Integration_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MSR_TRC_C----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call SetTextbox("Sales Document","SO_VBELN-LOW","",DT_MSR_TRC_C_0302_SALES_DOCUMENT,FALSE)
Call TakeScreenShot()

'Click on Yes If Pop Up Exist
Call ClickButtonIfExist("Yes",True) 
Wait(2)


'Click on Execute
Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(2)

Call ActivateNodeGuiTree(0,"Customer return with shipment to plant;Returns Delivery")
Call TakeScreenShot()

Call GetTextboxValue("LIKP-VBELN","0","DT_OUTBOUND_DELV_NO_OUTPUT",False)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

''----------------------Tcode vl02n----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MSR_TRC_C_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MSR_TRC_C_0100_OKCD)
Call TakeScreenShot()


'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_MSR_TRC_C_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Click on Post Goods Issue button
Call ClickButton("Back   \(F3\)",False)
'Call ClickButtonIfExist("Post Goods Receipt   \(Shift\+F8\)",False)
'wait(2)
Call ClickButtonIfExist("Post Goods Issue   \(Shift\+F8\)",False)
wait(2)
Call VerifyStatusBar(DT_MSR_TRC_C_4004_CHECK_TEXT_OF_STATUSBAR)



''----------------------Tcode vl03n----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MSR_TRC_C_4004_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MSR_TRC_C_4004_OKCD)
Call TakeScreenShot()


'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_MSR_TRC_C_4004_OUTBOUND_DELIVERY_OCC1,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Click on Document Flow
Call ClickButton("Document Flow   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

'Click on GUI Tree
Call ActivateItemGuiTree(0,"#1;#1;#1;#1;#1;#1","#1")
Wait(2)
Call TakeScreenShot()


'Get the Document No
Call GetGridContentByTitle("GD returns unrestr.*",0,"Doc.no.",1,"DT_DOCUMENT_NUM_OUTPUT")

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

''----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MSR_TRC_C_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MSR_TRC_C_0100_OKCD_OCC1)
Call TakeScreenShot()

Call SetCombo("GODYNPRO-ACTION","Display")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
Call SetComboByKey("GODYNPRO-REFDOC","R02")

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MSR_TRC_C_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,Year(Date),False)
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_MSR_TRC_C_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)

'Navigate to the Output tab
Call SelectTab("TS_GOITEM","Output",False)
Wait(1)
Call TakeScreenShot()

'Click on Display outputs
Call ClickButtonIfExist("Display outputs",False)
Wait(2)

'Verify the Output Type data
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_MSR_TRC_C_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(2)
Call TakeScreenShot()

Call DoubleClickGuiGridCell("Documents in Accounting",0,1,"Document Number",True)
Call TakeScreenShot()

Call  GetTextboxValue("BKPF-BELNR",0,"DT_ACCOUNTING_DOC_NO",False)

Call SelectColumnGuiGrid("",0,"Account",False)
Call ClickButtonToolBar("&MB_FILTER",0)

Call SetTextboxPopupIfExist("%%DYN001-LOW","Account",DT_MSR_TRC_C_1105_ACCOUNT)
Wait(1)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
Wait(2)


Call VerifyGridCellContent("",1,"Profit Center",0,DT_MSR_TRC_C_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)


'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Wait(2)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


