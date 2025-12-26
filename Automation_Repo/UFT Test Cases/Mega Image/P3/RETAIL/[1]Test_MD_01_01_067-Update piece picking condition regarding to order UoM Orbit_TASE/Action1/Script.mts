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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_MD_01_01_067-Update piece picking condition regarding to order UoM Orbit_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 15th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_MD_01_01_067-UoM Orbit_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_MD_01_01_067-Update piece picking condition regarding to order UoM Orbit_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter/reload
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode MM43----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,False)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM43_0100_SALES_ORG,False)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM43_0100_DISTR_CHANNEL,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()   
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTableCellContent(1,"AUoM","SAPLMGD2TC_ME_8022",DT_MM43_8022_CHECK_TEXT_OF_TABLECELL_AUOM_0)
Call VerifyTableCellContent(2,"AUoM","SAPLMGD2TC_ME_8022",DT_MM43_8022_CHECK_TEXT_OF_TABLECELL_AUOM_1)
''
''''----------------------Tcode zmdas_wsl11----------------------------
'Enter the Tcode
Call SetTcode(DT_MM43_4008_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MM43_4008_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_MM43_1000_ASSORTMENT,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_MM43_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",DT_MM43_1000_VALID_FROM,False)
Call SetTextbox("Valid To","P_DATBI","",DT_MM43_1000_VALID_TO,False)
Call SetTextbox("Customer No\. - Site","S_LOCNR-LOW","",DT_MM43_1000_SITE_FROM,False)

SAPGuiSession("transaction:=zmdas_wsl11").SAPGuiWindow("transaction:=zmdas_wsl11").SAPGuiButton("tooltip:=Multiple selection","index:=1").Click
Wait(2)
Call SelectTab("TAB_STRIP","Exclude Ranges",True)
Call SetTableDataNoRef("SAPLALDBINTERVAL_E","Lower limit",1,"R011",True)
Call SetTableDataNoRef("SAPLALDBINTERVAL_E","Upper limit",1,"R999",True)
Call SetTableDataNoRef("SAPLALDBINTERVAL_E","Lower limit",2,"RO00000000",True)
Call SetTableDataNoRef("SAPLALDBINTERVAL_E","Upper limit",2,"RO99999999",True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectColumnGuiGrid("","","Customer No. - Site",False)
Call ClickButtonIfExist("Sort in Ascending Order   \(Ctrl\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR)
Call VerifyGridCellContent("",2,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOCNR)
Call VerifyGridCellContent("",3,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LOCNR)
Call VerifyGridCellContent("",4,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_LOCNR)
Call VerifyGridCellContent("",5,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_LOCNR)
Call VerifyGridCellContent("",6,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_LOCNR)
Call VerifyGridCellContent("",7,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_LOCNR)
Call VerifyGridCellContent("",8,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_LOCNR)
Call VerifyGridCellContent("",9,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_LOCNR)
Call VerifyGridCellContent("",10,"Customer No. - Site","",DT_MM43_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_LOCNR)

''
''''----------------------Tcode VK11----------------------------
'Enter the Tcode
Call SetTcode(DT_MM43_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MM43_0500_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MM43_0100_CONDITION_TYPE,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Choose   \(Enter\)",True)

Call SetTextbox("Sales Organization","KOMG-VKORG","",DT_MM43_1969_SALES_ORGANIZATION,False)
Call SetTextbox("Article","KOMG-MATNR","",DT_MM43_1969_ARTICLE,False)

Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Site",1,DT_MM43_1969_TABLECELL_SITE_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","per",1,DT_MM43_1969_TABLECELL_PER_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","UoM",1,DT_MM43_1969_TABLECELL_UOM_0,False)

Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","Site",2,DT_MM43_1969_TABLECELL_SITE_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","per",2,DT_MM43_1969_TABLECELL_PER_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_FAST_ENTRY","UoM",2,DT_MM43_1969_TABLECELL_UOM_0,False)

Call PressEnter() 

'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Wait(2)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBar(DT_MM43_1969_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

