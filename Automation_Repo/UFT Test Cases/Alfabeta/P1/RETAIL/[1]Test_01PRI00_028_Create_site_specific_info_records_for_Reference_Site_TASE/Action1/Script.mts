

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_028_Create_site_specific_info_records_for_Reference_Site
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



gstrTestCaseName = "Test_01PRI00_028_Create_site_specific_info_records_for_Reference_Site"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_028_Create_site_specific_info_records_for_Reference_Site_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  MM42------------------------------------------------

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_100_ARTICLE,False)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_MM42_100_PURCHASING_ORG,False)
Call SetTextbox("Vendor Subrange","RMMW1-LTSNR","","",False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_100_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_100_DISTR_CHANNEL,False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Purchasing",False)
Call PressEnter() 
Call ClickButtonIfExist("Copy   \(Enter\)",True)

'Call ClickButtonIfExist("btn\[0\]",False)

Call TakeScreenShot()
Call GetTextboxValue("RMMW1-LIFNR",0,"DT_MM42_100_GET_TEXT_OF_VENDOR_OUTPUT",False)
Call PressEnter() 
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call  ClickButton("Switch area of validity   \(Shift\+F1\)",False)
Call SetTextbox("Site","RMMW1-EKWRK","",DT_MM42_81_SITE,True)
Call TakeScreenShot()
Call  ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call  ClickButton("For the Conditions",False)
Call  ClickButtonIfExist("Select time frame   \(Enter\)",True)
Call VerifyGridCellContent("",1,"LIFNR",0,DT_MM42_100_GET_TEXT_OF_VENDOR_OUTPUT)
Call VerifyGridCellContent("",1,"MATNR",0,DT_MM42_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",1,"EKORG",0,DT_MM42_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKORG)
Call VerifyGridCellContent("",1,"WERKS",0,DT_MM42_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WERKS)
Call GetTableCellData("SAPMV13ATCTRL_D0201","Amount",1,"","","DT_MM42_201_GET_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT",False)
Call SetTableData("SAPMV13ATCTRL_D0201","Amount",1,"","",DT_MM42_201_GET_TEXT_OF_TABLECELL_AMOUNT_NUM,False)
Call TakeScreenShot()
Call PressEnter() 
Call  ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()



'--------------------------------------------  ME21N------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_MM42_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_MM42_PO_TYPE)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_MM42_1105_VENDOR,False)   
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_MM42_1105_VENDOR,False)   

Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_MM42_1221_PURCH_ORG,False) 
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_MM42_1221_COMPANY_CODE,False)    
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_MM42_1211_TABLECELL_ARTICLE_0,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_MM42_1211_TABLECELL_PO_QUANTITY_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_MM42_1211_TABLECELL_SITE_0,False)  
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("ITEM_DETAIL","Conditions",False)
Call VerifyTableCellContent(1,"Name","SAPLV69ATCTRL_KONDITIONEN",UCASE(DT_MM42_6201_CHECK_TEXT_OF_TABLECELL_NAME_0))
Call PressEnter()
Call SelectTab("ITEM_DETAIL","Conditions",False)
'Call VerifyTableCellContent(1,"Amount","SAPLV69ATCTRL_KONDITIONEN",DT_MM42_201_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_MM42_6201_CHECK_TEXT_OF_TABLECELL_CNTY_0)

Call LogOff()
Call FinalStatus ()



