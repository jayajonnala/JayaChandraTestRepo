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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Self consumption for logicstic department in DC for DRY goods_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 14th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Self consumption for logicstic department in DC for DRY goods_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Self consumption for logicstic department in DC for DRY goods_p2_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login
'DataRowSet=2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VL10C----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call WriteRunTimeDataToExcelGlobalSheet ("DT_PO_INC",(Cint(DT_PO_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call SelectTab("TABSTRIP_ORDER_CRITERIA",DT_VL10C_1000_SALES_ORDERS,False)
Wait(1)


'Call SetTextbox("SD Document","ST_VBELN-LOW","",DT_VL10C_1020_SALES_DOCUMENT,False)
'Call SetTextbox("Sales Document","ST_VBELN-LOW","",DT_VL10C_1020_SALES_DOCUMENT,False)
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_VL10C_1020_SALES_DOCUMENT,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call SelectCheckboxNoLabel("0",DT_VL10C_0120_NO_NAME,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Create Delivery in Background   \(Shift\+F7\)",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

''Line 80 to 103 written according to TAO log, but after clicking log for delivery creation, a new tab is opened in SAP, as causing failure of script
''There fore, we use another Tcode, i.e VA03 in order to post goods issue and continue with VF01 as per TAO log





'Call SelectCheckboxNoLabel("1",DT_VL10C_0120_NO_NAME,False)
'Call ClickButtonIfExist("Log for delivery creation   \(Shift\+F4\)",False)
'Call VerifyStatusBarMessageType("S")
'
'Call SetFocusGuiLabel("AUTODS",81,56, False)
'Call ClickButton("Documents   \(F9\)",False)
'
'Call GetLabelContentByRefLabel("Group","77","32",DT_VL10C_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT,False)
'
'Call WriteRunTimeDataToExcel("DT_VL10C_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",DT_VL10C_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'Call SetFocusGuiLabel("",11,56, False)
'Call ClickButton("Display document   \(Shift\+F2\)",False)
'
'Call VerifyTextBoxContent("Outbound deliv\.","LIKP-VBELN","",DT_VL10C_0120_CHECK_TEXT_OF_NO_NAME_OCC1,False)
'Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)
'
'Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
'Call VerifyStatusBarMessageType("S")
'Call GetStatusBar("item1",DT_VL10C_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
'Call WriteRunTimeDataToExcel("DT_VL10C_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_VL10C_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusBar(DT_VL10C_0120_CHECK_TEXT_OF_STATUSBAR)
'




'''----------------------Tcode VA03-------alternative way to get the delivery number---------------------
Call SetTcode("VA03") 
Call PressEnter()  

Call SetTextbox("Order","VBAK-VBELN","",DT_VL10C_1020_SALES_DOCUMENT,False)' 
Call PressEnter()  
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call ClickButtonIfExist("Display document flow   \(F5\)",False)
Call ClickButtonToolBar("&FIND","")
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Open",True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Display document   \(F8\)",False)

Call GetTextboxValue("LIKP-VBELN","","DT_VL10C_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButtonIfExist("Display <-> Change   \(Ctrl\+F1\)",False)
Wait(2)
Call VerifyTextBoxContent("Outbound deliv\.","LIKP-VBELN","",DT_VL10C_1502_CHECK_TEXT_OF_OUTBOUND_DELIV,False)

Call ClickButtonIfExist("Post Goods Issue   \(Shift\+F8\)",False)
Wait(2)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_VL10C_0120_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item3","DT_ARTICLE_DOCUMENT_NUM_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_VL10C_0120_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
'
'''----------------------Tcode VF01----------------------------
Call SetTcode(DT_VL10C_0120_OKCD) 
Call PressEnter()
Call CheckTCodeScreen(DT_VL10C_0120_OKCD)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMV60ATCTRL_ERF_FAKT","Document",1,DT_VL10C_0102_TABLECELL_DOCUMENT_0,False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_VL10C_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_VL10C_0102_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

Call ClickButtonIfExist("Display   \(Ctrl\+F11\)",False)
Call SetTextbox("Billing document","VBRK-VBELN","",DT_VL10C_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR,False)' 
Call PressEnter() 
Call TakeScreenShot()

Call ClickButtonIfExist("Accounting overview   \(Shift\+F4\)",False)
Call TakeScreenShot()

''Call SelectCellGuiGrid("Documents in Accounting","",1,"Doc. Number",True)
Call SelectCellGuiGrid("Documents in Accounting","",1,"Document Number",True)
Call ClickButtonIfExist("Display Document   \(F2\)",False)
Wait(2)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


