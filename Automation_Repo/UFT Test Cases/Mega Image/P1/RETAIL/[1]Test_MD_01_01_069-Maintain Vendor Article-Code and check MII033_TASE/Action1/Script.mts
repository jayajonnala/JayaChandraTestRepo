
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_069-Maintain Vendor Article-Code and check MII033
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


gstrTestCaseName = "Test_MD_01_01_069-Maintain Vendor Article-Code and check MII033"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode---/nzmdpu_vend_art----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Purch. Organization","P_EKORG","",DT_ZMDPU_VEND_ART_1000_PURCH_ORGANIZATION,false)
''Call SetTextbox("Vendor","P_LIFNR","",DT_ZMDPU_VEND_ART_1000_VENDOR,false)
Call SetTextboxNoLabel("P_LIFNR","",DT_ZMDPU_VEND_ART_1000_VENDOR,false)
Call SetTextbox("Article","P_MATNR","",DT_ZMDPU_VEND_ART_1000_ARTICLE,false)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot

Call ClickButtonIfExist("Change\/Display   \(F7\)",False)

''''Call Click204ButtonToolBar("BBY_INS_ROW",0)
''''Call ClickButtonToolBar("DELE",0)
''''Call ClickButtonToolBar("WB2R_NEW_LINE",0)

SAPGuiSession("Session").SAPGuiWindow("Creation of multiple vendor's_2").InsightObject("InsightObject").Click



Wait 4
Call SetGridData("",1,"Vendor Article Number",DT_ZMDPU_VEND_ART_0100_GRIDCELL_0_VENDOR_ARTICLE_NUMBER,False)
Call SelectCheckBoxGridByRefColumn("shell","Primary Vendor Material Code - Index", "Vendor Article Number",DT_ZMDPU_VEND_ART_0100_GRIDCELL_0_VENDOR_ARTICLE_NUMBER,"ON")

Call ClickButtonToolBar("&DETAIL",0)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot

Call VerifyGridCellContentbyName("shell",1,"Vendor Article Number","",DT_ZMDPU_VEND_ART_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_IDNLF)

Call LogOff()

Call FinalStatus ()

