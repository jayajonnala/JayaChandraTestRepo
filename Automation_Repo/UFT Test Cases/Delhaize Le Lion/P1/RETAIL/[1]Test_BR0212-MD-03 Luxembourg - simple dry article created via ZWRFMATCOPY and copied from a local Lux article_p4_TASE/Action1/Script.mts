'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_BR0212-MD-03 Luxembourg - simple dry article created via ZWRFMATCOPY and copied from a local Lux article_p4_TASE
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



gstrTestCaseName = "Test_BR0212-MD-03 Luxembourg  Lux article_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\rtod\Documents\Input Datasheet\DLL\DT_04.04.02.12 VIM - NPO Precontrole Issue - BR10b - Invalid Vendor_TASE2.xls"



Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'DataRowSet =2

Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Wait(2)

'----------------------Tcode MM43----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,false)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM43_0100_SALES_ORG,false)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM43_0100_DISTR_CHANNEL,false)
Call TakeScreenShot()
Call ClickButtonIfExist("Deselect All   \(Shift\+F7\)",False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",1,False)
  
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("No",True)
Call TakeScreenShot()
Call VerifyTextBoxContent("X-plant status","MARA-MSTAE",0,DT_MM43_2004_CHECK_TEXT_OF_XSITE_STATUS,False)
Call VerifyTextBoxContent("Astmt List Type","MAW1-BBTYP",0,DT_MM43_2003_CHECK_TEXT_OF_ASTMT_LIST_TYPE,False)

Call LogOff()
Call FinalStatus ()

