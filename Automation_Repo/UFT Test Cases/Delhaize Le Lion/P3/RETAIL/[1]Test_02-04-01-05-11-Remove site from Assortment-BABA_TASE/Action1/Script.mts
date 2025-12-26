'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-11-Remove site from Assortment-BABA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-05-11-Re BABA"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-WSOA2----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Assortment","V_WRS1-ASORT","",DT_WSOA2_0001_ASSORTMENT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Basic Data   \(F8\)",False) 
Call TakeScreenShot
Call SelectTab("ASORT_TAB", "Assortment User", False)
Call TakeScreenShot
Call SelectRowGuiTableByRow("WRFM_WSO6DYN040",2,False)
Call ClickButtonIfExist("Delete",False)
Call VerifyTextBoxContent("Assortment","V_WRS1-ASORT","",lcase(DT_WSOA2_0010_CHECK_TEXT_OF_ASSORTMENT),False)
Call VerifyTableCellContent(1,"CustomerNoSite","WRFM_WSO6DYN040",DT_WSOA2_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_0)
Call VerifyTableCellContent(2,"CustomerNoSite","WRFM_WSO6DYN040",DT_WSOA2_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_1)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_WSOA2_0001_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
