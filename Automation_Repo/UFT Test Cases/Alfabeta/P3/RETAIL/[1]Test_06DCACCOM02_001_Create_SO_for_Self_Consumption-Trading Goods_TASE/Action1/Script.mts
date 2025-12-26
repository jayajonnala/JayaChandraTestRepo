

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06TRIS02SR_007_Create_for_TP_Returns_ZTRE
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

gstrTestCaseName = "Test_06TRIS02SR_007_ZTRE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

'''--------------------------------va01-----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Purch\. Order No\.","VBKD-BSTKD","",DT_VA01_4021_PO_NUMBER,False)
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call SetTextbox("Ship-To Party","KUWEV-KUNNR","",DT_VA01_4701_SHIPTO_PARTY,False)
Call SetTextbox("Delivering Site","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)
Call SetComboByKey("VBAK-AUGRU", DT_VA01_4440_ORDER_REASON)

Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article","1","","",DT_VA01_4900_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity","1","","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Storage Location","1","","",DT_VA01_4900_TABLECELL_STORAGE_LOCATION_0,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SelectMenuBar("Sales document;Deliver")
Call ClickButtonIfExist("Save",True) 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call GetStatusBar("item2","DT_SELF_CONS_NO_OUTPUT")

Call VerifyStatusBar("Self Consumption "&DT_SELF_CONS_NO_OUTPUT&" has been saved")
'Call WriteRunTimeDataToExcel ("DT_SELF_CONS_NO_OUTPUT",DT_SELF_CONS_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectTab("TAXI_TABSTRIP_OVERVIEW", "Picking", False)
Call TakeScreenShot()
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call TakeScreenShot()


Call GetStatusBar("item2","DT_SELF_CONS_GI_OUTPUT")
Call VerifyStatusBar("Self Consumption "&DT_SELF_CONS_GI_OUTPUT&" has been saved")
'Call WriteRunTimeDataToExcel ("DT_SELF_CONS_GI_OUTPUT",DT_SELF_CONS_GI)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Back   \(F3\)",False)
Wait 2
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
''--------------------------------VL06O-----------------------------
Call SetTcode(DT_VA01_0100_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document","1","","",DT_VA01_0102_TABLECELL_DOCUMENT_0,False)
Call ClickButton("Execute   \(F8\)",False)
CAll VerifyTableCellContent(1, "Article", "SAPMV60ATCTRL_UEB_FAKT", DT_VA01_0104_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
'Call WriteRunTimeDataToExcel ("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_VA01_0102_CHECK_TEXT_OF_STATUSBAR)

Call LogOff'
