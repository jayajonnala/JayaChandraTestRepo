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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Internal order master data maintenance_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 12th July
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Internal order master data maintenance_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\FICO\DT_Internal order master data maintenance_p1_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'increment
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode KO04----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickLinkGuiTree("#1","Order","",False)
Call SetTextbox("Order Type","G_SELFLD_TAB-LOW","",DT_KO04_0220_ORDER_TYPE,True)
Call SetTextbox("Order","G_SELFLD_TAB-LOW","",DT_KO04_0120_NO_NAME,True)
'Capture the screenshot
Call TakeScreenShot() 

Call ClickButton("Continue   \(Enter\)",True)
Wait(1)

Call SelectCheckboxNoLabel("1",DT_KO04_0120_NO_NAME_OCC2,True)
'Capture the screenshot
Call TakeScreenShot() 
Call ClickButton("Copy   \(Enter\)",True)

Call ClickLinkGuiTree("#1","Order","",False)
Call SetTextbox("Order Type","G_SELFLD_TAB-LOW","",DT_KO04_0220_ORDER_TYPE,True)
'Capture the screenshot
Call TakeScreenShot() 

Call ClickButton("Continue   \(Enter\)",True)
Wait(1)

Call SelectCheckboxNoLabel("2",DT_KO04_0120_NO_NAME_OCC2,True)
Call SelectCheckboxNoLabel("3",DT_KO04_0120_NO_NAME_OCC3,True)
Call SelectCheckboxNoLabel("4",DT_KO04_0120_NO_NAME_OCC4,True)
'Capture the screenshot
Call TakeScreenShot() 
Call ClickButton("Copy   \(Enter\)",True)

'Capture the screenshot
Call TakeScreenShot()

Call ClickCellGuiGrid("Personal worklist","","Order",1,"Order",DT_KO04_0120_NO_NAME,False)
'Capture the screenshot
Call TakeScreenShot() 

Call ClickButton("Change   \(Ctrl\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot() 

''
''''----------------------Tcode KO01----------------------------
''
'Enter the Tcode
Call SetTcode(DT_KO04_0100_OKCD) 
Call PressEnter()     ' 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Order Type","COAS-AUART","",DT_KO04_0100_ORDER_TYPE,False)
Call PressEnter()  

Call SetTextbox("Order","COAS-AUFNR","",DT_RESULT,False)
Call SetTextbox("Description","COAS-KTEXT","",DT_KO04_0600_DESCRIPTION,False)
Call SetTextbox("Company Code","COAS-BUKRS","",DT_KO04_0315_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP_600",DT_KO04_0600_CONTROL_DATA,False)
Call SetTextbox("Actual posted CCtr","COAS-CYCLE","",DT_KO04_0325_ACTUAL_POSTED_CCTR,False)
'Capture the screenshot
Call TakeScreenShot()


Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
VerifyStatusBar(DT_KO04_0100_CHECK_TEXT_OF_STATUSBAR)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

