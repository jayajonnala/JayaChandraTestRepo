'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_BR0212-MD-03 Luxembourg - simple dry article created via ZWRFMATCOPY and copied from a local Lux article_p3_TASE
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



gstrTestCaseName = "Test_BR0212-MD-03 Luxembourg Lux article_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\rtod\Documents\Input Datasheet\DLL\DT_04.04.02.12 VIM - NPO Precontrole Issue - BR10b - Invalid Vendor_TASE2.xls"



Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'DataRowSet =2

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Wait(2)

'----------------------Tcode MM42----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

'''''''''
Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",1,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",3,False)
Call PressEnter()
Call TakeScreenShot()

''''''''''''''''''
Call PressEnter()
Call ClickButton("Go to additional data   \(Ctrl\+F6\)",fALSE)
Call SelectTab("TABSPR1","Vendor Data",False)
Call TakeScreenShot()

'SAPGuiSession("Session").SAPGuiWindow("Change Article").InsightObject("InsightObject").Click

 @@ hightlight id_;_0_;_script infofile_;_ZIP::ssf1.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Change Article 7741511").InsightObject("InsightObject").Click @@ hightlight id_;_10_;_script infofile_;_ZIP::ssf2.xml_;_

Call SetGridData("",1,"Alternative Unit of Measure", DT_MM42_9901_GRIDCELL_0_AUN,False)
Call SetGridData("",1,"Packaging mat. type", DT_MM42_9901_GRIDCELL_0_PKMTT_OCC1,False)
Call SetGridData("",1,"Packaging Material", DT_MM42_9901_GRIDCELL_0_PACKAGING_MATERIAL_OCC1,False)
Call TakeScreenShot() 

 
Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)",fALSE)
Call ClickButton("Display item details",fALSE)

Call SelectCheckBoxTableByRefColumn("SAPLZMDAM_EXT_SCRTC_ECOMM_SUPP","Supply Chain Cle Art","UoM","CV","OFF")
Call SelectCheckBoxTableByRefColumn("SAPLZMDAM_EXT_SCRTC_ECOMM_SUPP","Supply Chain Cle Art","UoM","IP","ON")
Call TakeScreenShot()

Call SelectTab("INFOEXT","Sales 1",False)
Call TakeScreenShot()
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Consumption Units Per Article",1,DT_MM42_2008_TABLECELL_CONSUMPTION_UNITS_PER_ARTICLE_0,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Net Content Per Consumption Unit",1,DT_MM42_2008_TABLECELL_NET_CONTENT_PER_CONSUMPTION_UNIT_0,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Content Unit Consump",1,DT_MM42_2008_TABLECELL_CONTENT_UNIT_CONSUMP_0,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Consumption Units Per Article",2,DT_MM42_2008_TABLECELL_CONSUMPTION_UNITS_PER_ARTICLE_1,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Net Content Per Consumption Unit",2,DT_MM42_2008_TABLECELL_NET_CONTENT_PER_CONSUMPTION_UNIT_1,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Content Unit Consump",2,DT_MM42_2008_TABLECELL_CONTENT_UNIT_CONSUMP_1,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Consumption Units Per Article",3,DT_MM42_2008_TABLECELL_CONSUMPTION_UNITS_PER_ARTICLE_2,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Net Content Per Consumption Unit",3,DT_MM42_2008_TABLECELL_NET_CONTENT_PER_CONSUMPTION_UNIT_2,False)
Call SetTableDataNoRef("SAPLZMDAM_EXT_SCRTC_ECOMM","Content Unit Consump",3,DT_MM42_2008_TABLECELL_CONTENT_UNIT_CONSUMP_2,False)
Call PressEnter()
Call TakeScreenShot()


Call SelectTab("INFOEXT","Logistics: DC/Store",False)
Call TakeScreenShot()
Call SetTextbox("Lifespan in Store - Manual","ZMDAM_MARAEXT-INSTOREMAN","",DT_MM42_2003_STORE,False)
Call TakeScreenShot()

Call SelectTab("INFOEXT","Store",False)
Call TakeScreenShot()
Call SetTextbox("Label type","ZMDAM_MARAEXT-LABEL_TYPE","",DT_MM42_2006_ZMDAM_MARAEXTLABEL_TYPE,False)
Call TakeScreenShot()
Call FocusTextBox("Label type","ZMDAM_MARAEXT-LABEL_TYPE",False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",fALSE)
Call ClickButton("Save   \(Ctrl\+S\)",fALSE)
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",1,False)
Call PressEnter()
''Call SetTextbox("X-site status","MARA-MSTAE","",DT_MM42_2004_XSITE_STATUS,False)
'Call SetTextboxNoLabel("MARA-MSTAE","",DT_MM42_2004_XSITE_STATUS,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",fALSE)
Wait 4
Call PressEnter()
'Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


