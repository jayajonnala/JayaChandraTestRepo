
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p1
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
'.................Test Script Name :Test_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    :4th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode ME21N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call PressEnter()

'Enter the Organisation details
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

'Enter Order Details
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","","",False) 

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
Call PressEnter()
Call PressEnter()

'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Wait(2)
Call ClickButtonIfExist("Save",True)

'Validate If STO Retail Order is generated
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
VerifyStatusBar("Standard PO Retail created under the number " & DT_PO_NUMBER_OUTPUT)



'----------------------Tcode ME29N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_0014_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_ME21N_0014_OKCD)

'Click on Other purchase order
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(1)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME21N_0003_PUR_ORDER)
Call TakeScreenShot()
Call ClickButtonIfExist("Other Document   \(Enter\)",True)
wait(2)
Call TakeScreenShot()

'Expand the Header
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot()
wait(2)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call TakeScreenShot()


'Navigate to the Status Tab
Call SelectTab("HEADER_DETAIL","Status",False)
Wait(1)
Call TakeScreenShot()


Call GetTextboxValue("MEPO1234-QUANTITY01",0,"TEXT_OF_ORDERED_QTY",False)
Call GetTextboxValue("MEPO1235-VALUE01",0,"TEXT_OF_RSD_OUTPUT",False)

'Click on Save
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(1)
Call TakeScreenShot()


'Call VerifyTextPopup(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC2)

''Verify Text Box Content
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,UCASE(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC2),True)


'Click on Continue button
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(1)


''----------------------Tcode ME23N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME21N_0014_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME21N_0014_OKCD_OCC1)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME21N_0003_PUR_ORDER_OCC1)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

'Click on Messages button and get the output type.
Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
wait(2)
Call TakeScreenShot()



'Verify the Status
Call VerifyTableCellContent(3,"Status","SAPDV70ATC_NAST3",DT_ME21N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_2)
Call VerifyTableCellContent(4,"Status","SAPDV70ATC_NAST3",DT_ME21N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_3)

'Verify the Output Type data
Call VerifyTableCellContent(3,"Output Type","SAPDV70ATC_NAST3",DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)
Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)


''----------------------Tcode ME9F----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME21N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME21N_0100_OKCD)

'Enter the details
Call SetTextbox("Document Number","S_EBELN-LOW","",DT_ME21N_1000_DOCUMENT_NUMBER,False)   
Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME21N_1000_PURCHASING_ORGANIZATION,False)   
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_ME21N_1000_PURCHASING_GROUP,False)   
Call SetTextbox("Application","P_KAPPL","",DT_ME21N_1000_APPLICATION,False)   
Call SetTextbox("Processing Status","P_VSTAT","",DT_ME21N_1000_PROCESSING_STATUS,False)  
Call TakeScreenShot()

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'Select CheckBox
Call SelectCheckBoxByGuiLabel("ZNAB",38,1,DT_ME21N_0120_NO_NAME)
Call TakeScreenShot()
Call ClickButton("Trial Printout   \(Shift\+F4\)",False) 
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

'Click on Back
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

