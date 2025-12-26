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

''Reload DataSheet to updates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Park FI documents_p5_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 11th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Park FI documents_p5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Park FI documents_p5_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FBV0----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Create Document List   \(F5\)",False)
Wait(2)
Call SetTextbox("Company code","BUKRS-LOW","",DT_FBV0_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal year","GJAHR-LOW","",DT_FBV0_1000_FISCAL_YEAR,False)
Call SetTextbox("Document date","BLDAT-LOW","",Replace((DT_FBV0_1000_DOCUMENT_DATE),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectColumnGuiGrid("","",DT_FBV0_0600_GRIDCELL_4_COLUMN_NAME,False)
Wait(1)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+F5\)",False)
Wait(1)
Call ClickButtonIfExist("Multiple selection",True)
Wait(1)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FBV0_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_FBV0_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_FBV0_3010_TABLECELL_SINGLE_VALUE_2,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Copy   \(F8\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,DT_FBV0_0600_GRIDCELL_4_COLUMN_NAME,"",DT_FBV0_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",2,DT_FBV0_0600_GRIDCELL_4_COLUMN_NAME,"",DT_FBV0_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("",3,DT_FBV0_0600_GRIDCELL_4_COLUMN_NAME,"",DT_FBV0_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
'Capture the screenshot
Call TakeScreenShot()

Call DoubleClickGuiGridCell("","",1,DT_FBV0_0600_GRIDCELL_4_COLUMN_NAME,False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Document Number","ACGL_HEAD-BELNR","",DT_FBV0_1010_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)

Call ClickButtonIfExist("Post   \(Shift\+F11\)",False)
Wait(2)

Call VerifyTextBoxContent("Information Message","MESSTXT1","",Lcase(DT_FBV0_0010_CHECK_TEXT_OF_MESSTXT1),True)

Call ClickButtonIfExist("Continue   \(Enter\)",True)
'
'''----------------------Tcode FBV3----------------------------
'Enter the Tcode
Call SetTcode(DT_FBV0_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FBV0_0500_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Create Document List   \(F5\)",False)
Wait(2)
Call SetTextbox("Company code","BUKRS-LOW","",DT_FBV0_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal year","GJAHR-LOW","",DT_FBV0_1000_FISCAL_YEAR_OCC1,False)
Call SetTextbox("Document date","BLDAT-LOW","",Replace((DT_FBV0_1000_DOCUMENT_DATE_OCC1),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

