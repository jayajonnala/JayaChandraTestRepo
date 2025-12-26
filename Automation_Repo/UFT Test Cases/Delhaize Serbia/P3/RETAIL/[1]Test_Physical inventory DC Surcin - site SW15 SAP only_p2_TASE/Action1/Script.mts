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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Physical inventory DC Surcin - site SW15 SAP only_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 7th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Physical inventory DC Surcin - site SW15 SAP only_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Physical inventory DC Surcin - site SW15 SAP only_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
''Increment the parameter
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_MI31_0731_ISEGERFMG",(Cint(DT_MI31_0731_ISEGERFMG)+1))
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode MI07----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Phys\. Inventory Doc\.","RM07I-IBLNR","",DT_MI07_0701_PHYS_INVENTORY_DOC,False)
Call SetTextbox("Fiscal Year","RM07I-GJAHR","",DT_MI07_0701_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(3)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item2","DT_MI07_0701_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_MI07_0701_CHECK_TEXT_OF_STATUSBAR)
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_MI07_0701_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MI07_0701_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetCombo("GODYNPRO-ACTION",DT_MI07_0010_GODYNPROACTION)
'Call SetCombo("GODYNPRO-REFDOC",DT_MI07_0010_GODYNPROACTION_OCC1)
Call SetComboByKey("GODYNPRO-REFDOC","R02")
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MI07_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MI07_0701_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD","Doc. info",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("FI Documents",False)
'Capture the screenshot
Call TakeScreenShot()

'Call GetTextboxValue("BKPF-BELNR","","DT_MI07_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

