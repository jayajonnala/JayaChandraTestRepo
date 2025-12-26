'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_BR0212-MD-03 Luxembourg - simple dry article created via ZWRFMATCOPY and copied from a local Lux article_p2_TASE
'.................Author : TCS_Ramesh
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_BR0212-MD-03 Luxembourg  Lux article_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\rtod\Documents\Input Datasheet\DLL\DT_04.04.02.12 VIM - NPO Precontrole Issue - BR10b - Invalid Vendor_TASE2.xls"



Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'DataRowSet =2
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Wait(2)

''----------------------Tcode VKP5----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call TakeScreenShot()
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'''''''''
Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",fALSE)
Call SetTextbox("Created By","ENAME-LOW","",FormatBlank(DT_VKP5_0100_CREATED_BY),True)
Call ClickButton("Execute   \(F8\)",True) 

Call SelectRowGuiGrid("Variant Catalog .*","","Variant name",DT_VKP5_0841_SEARCH_TERM,True)
Call TakeScreenShot()

Call ClickButton("Choose   \(F2\)",True)
Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)

Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","S_PLTYP-LOW","",DT_VKP5_1000_PRICE_LIST,False)
Call SetTextbox("to","S_PLTYP-HIGH","",DT_VKP5_1000_TO,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDate(DT_VKP5_1000_VALIDITY),False)
Call TakeScreenShot()

Call PressEnter()
Call PressEnter()
Call ClickButton("Execute   \(F8\)",fALSE)
Call PressEnter()
Call ClickButton("Select All   \(F7\)",fALSE)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",fALSE)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_Output")
'Call VerifyStatusBar("Data saved; pricing document "&DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_Output&" created")

'----------------------Tcode VK11----------------------------

Call SetTcode(DT_VKP5_1000_OKCD)
Call PressEnter()
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VKP5_0100_CONDITION_TYPE,False)
Call TakeScreenShot()
 
Call PressEnter()
Call SelectRadioButton("RV130-SELKZ","Sales Org\./Dist\. Channel/Price List/Article/Sales Unit",True)
Call TakeScreenShot()

Call PressEnter()
Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_VKP5_1155_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","KOMG-VTWEG","",DT_VKP5_1155_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Price List","KOMG-PLTYP","",DT_VKP5_1155_PRICE_LIST,False)
Call FocusTextBox("Price List","KOMG-PLTYP", False)
Call PressEnter()
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_VKP5_1155_TABLECELL_ARTICLE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",1,DT_VKP5_1155_TABLECELL_SALES_UNIT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_VKP5_1155_TABLECELL_AMOUNT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",2,DT_VKP5_1155_TABLECELL_ARTICLE_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",2,DT_VKP5_1155_TABLECELL_SALES_UNIT_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",2,DT_VKP5_1155_TABLECELL_AMOUNT_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",3,DT_VKP5_1155_TABLECELL_ARTICLE_2,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",3,DT_VKP5_1155_TABLECELL_SALES_UNIT_2,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",3,DT_VKP5_1155_TABLECELL_AMOUNT_2,False)
Call PressEnter()
Call TakeScreenShot()


Call ClickButton("Save   \(Ctrl\+S\)",fALSE)
Call VerifyStatusBar(DT_VKP5_1155_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

Call SetTextbox("Price List","KOMG-PLTYP","",DT_VKP5_1155_PRICE_LIST_OCC1,False)
Call FocusTextBox("Price List","KOMG-PLTYP", False)
Call PressEnter()
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",1,DT_VKP5_1155_TABLECELL_ARTICLE_0_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",1,DT_VKP5_1155_TABLECELL_SALES_UNIT_0_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,DT_VKP5_1155_TABLECELL_AMOUNT_0_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",2,DT_VKP5_1155_TABLECELL_ARTICLE_1_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",2,DT_VKP5_1155_TABLECELL_SALES_UNIT_1_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",2,DT_VKP5_1155_TABLECELL_AMOUNT_1_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Article",3,DT_VKP5_1155_TABLECELL_ARTICLE_2_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Sales unit",3,DT_VKP5_1155_TABLECELL_SALES_UNIT_2_OCC1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Amount",3,DT_VKP5_1155_TABLECELL_AMOUNT_2_OCC1,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",fALSE)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call VerifyStatusBar(DT_VKP5_1155_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

