
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_MD_01_01_069-Maintain Vendor Article-Code and check MII033_Delete lines
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

gstrTestCaseName = "Test_POST_MD_01_01_069-Maintain Vendor Article-Code and check MII033_Delete lines"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''--------------ZMDPU_VEND_ART----------''''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Purch. Organization","P_EKORG","",DT_ZMDPU_VEND_ART_1000_PURCH_ORGANIZATION,False)
''Call SetTextbox("Vendor","P_LIFNR","",DT_ZMDPU_VEND_ART_1000_VENDOR,False)
Call SetTextboxNoLabel("P_LIFNR","",DT_ZMDPU_VEND_ART_1000_VENDOR,False)
Call SetTextbox("Article","P_MATNR","",DT_ZMDPU_VEND_ART_1000_ARTICLE,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("Change\/Display   \(F7\)",False)

''''-----Delete Articles----'''''

Call SelectAllRowGuiGrid("",0,False)
SAPGuiSession("Session").SAPGuiWindow("Creation of multiple vendor's").InsightObject("InsightObject").Click



Call ClickButtonToolBar("&DETAIL",0)
Call PressEnter()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call PressEnter()

Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot
Call verifyNoRowExistsGrid("shell", 0)


Call LogOff()

Call FinalStatus ()



