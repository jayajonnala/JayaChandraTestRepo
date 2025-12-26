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
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_S2A_PRI_03_16-Create sales order to domestic partner_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_S2A_PRI_03_16-Create sales order to domestic partner_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_S2A_PRI_03_16-Create sales order to domestic partner_TASE.xls"
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
''''----------------------Tcode VK13----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VK13_0100_CONDITION_TYPE,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButtonIfPopupExists("RV130-SELKZ",DT_VK13_0100_SALES_ORGDISTR_CHL)
Call ClickButtonIfExist("Choose   \(Enter\)",True)

Call SetTextbox("Sales Organization","F001","",DT_VK13_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","F002-LOW","",DT_VK13_1000_DISTRIBUTION_CHANNEL,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"Distr. Chl",DT_VK13_1000_DISTRIBUTION_CHANNEL,"DT_VK13_1616_CHECK_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''
''''----------------------Tcode VA01----------------------------
'Enter the Tcode
Call SetTcode(DT_VK13_1616_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VK13_1616_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Order Type","VBAK-AUART","",DT_VK13_0101_ORDER_TYPE,False)
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VK13_0101_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VK13_0101_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Division","VBAK-SPART","",DT_VK13_0101_DIVISION,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  

Call SetTextbox("Purch\. Order No\.","VBKD-BSTKD","",DT_VK13_4021_PO_NUMBER,False)
Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VK13_4701_SOLDTO_PARTY,False)
Call PressEnter() 

Call SetTextbox("Incoterms","VBKD-INCO1","",DT_VK13_4440_INCOTERMS,False)
Call SetTextbox("Incoterms","VBKD-INCO2","",DT_VK13_4440_INCOTERMS_OCC1,False)
Call PressEnter() 

Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article",1,DT_VK13_4900_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity",1,DT_VK13_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call SetTableDataNoRef("SAPMV45ATCTRL_U_ERF_AUFTRAG","Site",1,DT_VK13_4900_TABLECELL_SITE_0,False)
Call PressEnter()
Call PressEnter()
Call SelectCellGuiTable("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article","Item","10",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Item conditions",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTableCellContent(2,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_VK13_6201_CHECK_TEXT_OF_TABLECELL_CNTY_1)
Call VerifyTableCellContent(2,"Amount","SAPLV69ATCTRL_KONDITIONEN",DT_VK13_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

