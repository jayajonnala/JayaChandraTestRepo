'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-03-Create new assortment-BASA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-05-03-CrBASA"
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

''''''--------TransactionCode-WSOA1----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSORTMENT_UNIQUE",Cint(DT_ASSORTMENT_UNIQUE)+1)
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Assortment","V_WRS1-ASORT","",DT_WSOA1_0001_ASSORTMENT,False)
Call SetTextbox("Assortment Type","WRS1-ASSORTYP","",DT_WSOA1_0001_ASSORTMENT_TYPE,False)
Call PressEnter() 
Call TakeScreenShot
Call SetTextbox("Assortment","V_WRS1-NAME1","",DT_WSOA1_0010_ASSORTMENT,False)
Call TakeScreenShot
Call SetTextbox("Sales Organization","V_WRS1-VKORG","",DT_WSOA1_0015_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","V_WRS1-VTWEG","",DT_WSOA1_0015_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("LEGACYASSORTMENT NR","WRS1-ASSORDIMVAL1","",DT_WSOA1_0015_LEGACYASSORTMENT_NR,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("ASORT_TAB", "Langs", False)
Call TakeScreenShot
Call PressEnter()
Call SetTableData("WRFM_WSO6DYN020", "Lang.", 2, "", "", DT_WSOA1_0020_TABLECELL_LANG_1, false)
Call SetTableData("WRFM_WSO6DYN020", "Description", 2, "", "", DT_WSOA1_0020_TABLECELL_DESCRIPTION_1, false)
Call SetTableData("WRFM_WSO6DYN020", "Lang.", 3, "", "", DT_WSOA1_0020_TABLECELL_LANG_2, false)
Call SetTableData("WRFM_WSO6DYN020", "Description", 3, "", "", DT_WSOA1_0020_TABLECELL_DESCRIPTION_2, false)
Call SetTableData("WRFM_WSO6DYN020", "Lang.", 4, "", "", DT_WSOA1_0020_TABLECELL_LANG_3, false)
Call SetTableData("WRFM_WSO6DYN020", "Description", 4, "", "", DT_WSOA1_0020_TABLECELL_DESCRIPTION_3, false)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("ASORT_TAB", "Assortment User", False)
Call SetTableData("WRFM_WSO6DYN040", "CustomerNoSite", 1, "", "", DT_WSOA1_0040_TABLECELL_CUSTOMERNOSITE_0, false)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Create Article Segments Using a Separate Job   \(Ctrl\+F1\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Immediate",True)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_WSOA1_0001_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
