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
'''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_S2A_PRI_02_049-Compare PBandNB prices_Check Prices_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 16th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_S2A_PRI_02_049-Compare PBandNB prices_Check Prices_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_S2A_PRI_02_049-Compare PBandNB prices_Check Prices_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter/reload
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
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
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Condition Information   \(Shift\+F4\)",False)

Call SetTextbox("Sales Organization","F002-LOW","",DT_VK13_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","F003-LOW","",DT_VK13_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Article","F001-LOW","",DT_VK13_1000_ARTICLE,False)
Call SetTextbox("from / on","SEL_DATE","",DT_VK13_1000_FROM__ON,False)

Call FocusTextBox("Price List","F007-LOW",False)
Wait(1)
Call SendKey("{F2}")
Wait(3)
Call SelectRowGuiGridbyRowNo("Price List","",1,True)
Call SelectCellGuiGrid("Price List","",1,"Description",True)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
Wait(1)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel("1",DT_VK13_0120_NO_NAME,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Display   \(F5\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_D0201",DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","F001-LOW","",DT_VK13_1000_ARTICLE_OCC1,False)
Call SetTextbox("from / on","SEL_DATE","",DT_VK13_1000_FROM__ON_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call SelectCheckboxNoLabel("1",DT_VK13_0120_NO_NAME_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Display   \(F5\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_D0201",DT_VK13_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0_OCC1)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

