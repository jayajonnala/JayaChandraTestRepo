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

'reload DS to update dates and calculations
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Maintain Asset Master Data_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Maintain Asset Master Data_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Maintain Asset Master Data_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''Reload
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Master data   \(F7\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False) 
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode AS02----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS03_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS03_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Master data   \(F7\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS03_1145_COST_CENTER,False)
Call SetTextbox("Location","ANLZ-STORT","",DT_AS03_1145_LOCATION,False)
Call SetTextbox("Profit Center","ANLZ-PRCTR","",DT_AS03_1145_PROFIT_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Yes   \(Enter\)",True) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

'Validate If asset is created
Call GetStatusBar("item1","DT_AS03_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS03_0100_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

