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
  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_S2A_PRI_03_19-Create sales order for export in RON_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_S2A_PRI_03_19-Create sales order for export in RON_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_S2A_PRI_03_19-Create sales order for export in RON_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter/reload
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode VA01----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  

'Call SetTextbox("Purch\. Order No\.","VBKD-BSTKD","",DT_VK13_4021_PO_NUMBER,False)
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False)
Call PressEnter() 

Call SetTextbox("Delivering Site ","RV45A-DWERK","",DT_VA01_4440_DELIVERING_SITE,False)

Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",1,DT_VA01_4900_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",2,DT_VA01_4900_TABLECELL_ARTICLE_1,False)

Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",1,DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",2,DT_VA01_4900_TABLECELL_ORDER_QUANTITY_1,False)

Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Site",1,DT_VA01_4900_TABLECELL_SITE_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Site",1,DT_VA01_4900_TABLECELL_SITE_1,False)
Call PressEnter()
Wait(2)
Call PressEnter()
Wait(2)
Call PressEnter()
Wait(2)

Call SelectRowGuiTable("SAPMV45ATCTRL_U_ERF_AUFTRAG","Item","10",False)
Call SelectRowGuiTable("SAPMV45ATCTRL_U_ERF_AUFTRAG","Item","20",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Edit;Fast change of...;Delivery date..")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F7\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)

Call ClickButton("Display header details",False)
'Capture the screenshot
Call TakeScreenShot()


Call SetTextbox("Doc\. Currency","VBAK-WAERK","",DT_VA01_4301_DOC_CURRENCY,False)
Call PressEnter()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call VerifyTextBoxContent("/","VBKD-KURSK","",DT_VA01_4301_CHECK_TEXT_OF_VBKDKURSK,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextStatusBar("DT_VA01_4001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

