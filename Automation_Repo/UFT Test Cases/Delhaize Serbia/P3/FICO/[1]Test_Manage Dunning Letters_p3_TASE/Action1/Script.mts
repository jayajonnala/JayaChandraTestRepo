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
'.................Test Script Name : Test_Manage Dunning Letters_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 5th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Dunning Letters_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Dunning Letters_p3_TASE.xls"
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
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode XD02----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("","RF02D-KUNNR","",DT_XD02_7101_RF02DKUNNR,True)
Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD02_7101_COMPANY_CODE,True)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Company code","RF02D-BUKRS",True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
Call ClickButtonIfExist("Company Code Data   \(Ctrl\+F2\)",False)
Call SelectTab("TABSTRIP100","Correspondence",False)
Call SetTextbox("Dunn\.Procedure","KNB5-MAHNA","",DT_XD02_7220_DUNNPROCEDURE,False)
Call SetTextbox("Last Dunned","KNB5-MADAT","","",False)'set blank : DT_DUNNING_LEVEL
Call SetTextbox("Dunning Level","KNB5-MAHNS","","",False)'set blank: DT_LAST_DUNNED
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

'save details
Call ClickButton("Save   \(Ctrl\+S\)",False)
'close the pop-up
Call ClickButtonIfExist("Cancel   \(F12\)",True)

'
'''----------------------Tcode F150----------------------------
'Enter the Tcode
Call SetTcode(DT_XD02_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_XD02_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Run On","F150V-LAUFD","",Replace((DT_XD02_0100_RUN_ON),"/","."),False)
Call SetTextbox("Identification","F150V-LAUFI","",DT_XD02_0100_IDENTIFICATION,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Identification","F150V-LAUFI",False)
Call GetTextboxValue("F150V-LAUFD","","DT_DATE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("DU_TABSTRIP",DT_XD02_0100_PARAMETER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Dunning date","F150V-AUSDT","",Replace((DT_XD02_0111_DUNNING_DATE),"/","."),False)
Call SetTextbox("Docmnts posted up to","F150V-GRDAT","",Replace((DT_XD02_0111_DOCMNTS_POSTED_UP_TO),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","RNG_BKRS-LOW","",DT_XD02_0001_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","RNG_SELC-LOW","",DT_XD02_0002_CUSTOMER,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Customer","RNG_SELC-LOW",False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("DU_TABSTRIP",DT_XD02_0100_FREE_SELECTION,False)
'Capture the screenshot
Call TakeScreenShot()

Call SendKey("{TAB}")
Wait(2)
Call SendKey("{TAB}")
Wait(2)
Call SendKey(DT_XD02_0112_TABLECELL_VALUES_0)
Wait(3)
Call SendKey("{TAB}")
Wait(3)
Call SendKey(DT_XD02_0112_TABLECELL_VALUES_0_OCC1)
Wait(3)
Call SetTableDataNoRef("SAPF150VDU_FLDTAB","Field name",1,DT_DOCUMENTS,False)
Call SetTableDataNoRef("SAPF150VDU_FLDTAB","Exclude values",1,DT_XD02_0112_TABLECELL_EXCLUDE_VALUES_0,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(3)
VerifyStatusBar(DT_XD02_0100_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("DU_TABSTRIP",DT_XD02_0100_STATUS,False)
'Capture the screenshot
Call TakeScreenShot()

'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxNoLabelContent("F150V-STEXT","",DT_XD02_0110_CHECK_TEXT_OF_F150VSTEXT,False)
Call ClickButton("Schedule dunning run   \(F7\)",False)
Wait(3)
Call SetTextbox("Output Device","USR01-SPLD","",DT_XD02_1100_OUTPUT_DEVICE,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("F150V-XSTRF","1",DT_XD02_1000_START_IMMEDIATELY,True)
Call SelectCheckbox("F150V-XDRAN","1",DT_XD02_1000_DUNNPRINT_WITH_SCHEDULING_,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F5\)",True)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

VerifyStatusBar(DT_XD02_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
'Capture the screenshot
Call TakeScreenShot()
Wait(10)
Call GetTextboxValue("F150V-STEXT","3","DT_TEXT_OUTPUT",False)
Call PressEnter() 
wait(2)
Call GetTextboxValue("F150V-STEXT","3","DT_TEXT_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot()
Call PressEnter() 
wait(2)
Call TakeScreenShot()

'Call VerifyTextBoxNoLabelContent("F150V-STEXT","3",DT_XD02_0110_CHECK_TEXT_OF_F150VSTEXT_OCC1,False)
'Call VerifyTextBoxNoLabelContent("F150V-STEXT","4",DT_XD02_0110_CHECK_TEXT_OF_F150VSTEXT_OCC2,False)

Call ClickButton("Exit   \(Shift\+F3\)",False)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

