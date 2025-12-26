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

gstrTestCaseName = "Test_Maintain Asset Master Data_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Maintain Asset Master Data_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Maintain Asset Master Data_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''Reload
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS02----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE,False)
'Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Master data   \(F7\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Deprec. Areas",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLAISTTC_ANLB","UseLife",5,DT_AS02_1190_TABLECELL_USELIFE_4,False)
Call SelectCellGuiTable("SAPLAISTTC_ANLB","UseLife","Area number","30",False)'substitue of focus cell of row no 5
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
Call GetStatusBar("item1","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS02_0100_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode AS05----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS02_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS02_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Block   \(Ctrl\+Shift\+F2\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("ANLA-XSPEB",DT_AS02_0200_LOCKED_TO_ACQUIS,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS02_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call ClickButton("Exit   \(Shift\+F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode AS06----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS02_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS02_0100_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET_OCC2,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Delete   \(Ctrl\+Shift\+F12\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RA02S-XPHYS",DT_AS02_0210_PHSYICALLY_DELETE_ASSET,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Yes",True)
''Wait(1)
'''Capture the screenshot
''Call TakeScreenShot()

'Validate If asset is created
Call GetStatusBar("item1","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_AS02_0100_CHECK_TEXT_OF_STATUSBAR_OCC2)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

