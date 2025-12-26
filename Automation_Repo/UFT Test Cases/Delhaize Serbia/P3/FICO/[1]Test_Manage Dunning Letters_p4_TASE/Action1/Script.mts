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
'.................Test Script Name : Test_Manage Dunning Letters_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 6th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Dunning Letters_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Dunning Letters_p4_TASE.xls"
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

''''----------------------Tcode F150----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Run On","F150V-LAUFD","",Replace((DT_F150_0100_RUN_ON),"/","."),False)
Call SetTextbox("Identification","F150V-LAUFI","",DT_F150_0100_IDENTIFICATION,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Identification","F150V-LAUFI",False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Schedule sample dunn\.printout\.\.\.   \(F8\)",False)
Wait(3)

Call SetTextbox("Output Device","USR01-SPLD","",DT_F150_1100_OUTPUT_DEVICE,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("F150V-XSTRF","1",DT_F150_1000_START_IMMEDIATELY,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Print   \(Ctrl\+P\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

VerifyStatusBar(DT_F150_0100_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()
Wait(10)
Call PressEnter() 
wait(2)
Call TakeScreenShot()

'Call VerifyTextBoxNoLabelContent("F150V-STEXT","6",DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT,False)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()


''----------------------Tcode SP02----------------------------
'Enter the Tcode
Call SetTcode(DT_F150_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_F150_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call ClickLabel("Date","",False)
Call ClickButton("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)
Call TakeScreenShot()

Call ClickLabel("OTFDOC","0",False)
Wait(2)
Call SelectMenuBar("Goto;List Display")
'Call VerifyifGuiLabelExists(DT_F150_0120_CHECK_TEXT_OF_1_PISMO_OPOMENE)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButton("Exit   \(Shift\+F3\)",False)
''
''''----------------------Tcode F150----------------------------
'Enter the Tcode
Call SetTcode(DT_F150_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_F150_0100_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Run On","F150V-LAUFD","",Replace((DT_F150_0100_RUN_ON_OCC1),"/","."),False)
Call SetTextbox("Identification","F150V-LAUFI","",DT_F150_0100_IDENTIFICATION_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Identification","F150V-LAUFI",False)
Call GetTextboxValue("F150V-LAUFD","","DT_DATE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("DU_TABSTRIP",DT_F150_0100_PARAMETER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Dunning date","F150V-AUSDT","",Replace((DT_F150_0111_DUNNING_DATE),"/","."),False)
Call SetTextbox("Docmnts posted up to","F150V-GRDAT","",Replace((DT_F150_0111_DOCMNTS_POSTED_UP_TO),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","RNG_BKRS-LOW","",DT_F150_0001_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","RNG_SELC-LOW","",DT_F150_0002_CUSTOMER,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Customer","RNG_SELC-LOW",False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("DU_TABSTRIP",DT_F150_0100_FREE_SELECTION,False)
'Capture the screenshot
Call TakeScreenShot()

Call SendKey("{TAB}")
Wait(2)
Call SendKey("{TAB}")
Wait(2)
Call SendKey(DT_F150_0112_TABLECELL_VALUES_0)
Wait(3)
Call SendKey("{TAB}")
Wait(3)
Call SendKey(DT_F150_0112_TABLECELL_VALUES_0_OCC1)
Wait(3)
Call SetTableDataNoRef("SAPF150VDU_FLDTAB","Field name",1,DT_F150_0112_TABLECELL_FIELD_NAME_0,False)
Call SetTableDataNoRef("SAPF150VDU_FLDTAB","Exclude values",1,DT_EXCLUDE_VALUES,False)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(3)
VerifyStatusBar(DT_F150_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("DU_TABSTRIP",DT_F150_0100_STATUS,False)
'Capture the screenshot
Call TakeScreenShot()

'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxNoLabelContent("F150V-STEXT","",DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC1,False)
Call ClickButton("Schedule dunning run   \(F7\)",False)
Wait(3)
Call SetTextbox("Output Device","USR01-SPLD","",DT_F150_1100_OUTPUT_DEVICE_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Start date","F150V-STRDT","",Replace((DT_F150_1000_START_DATE),"/","."),True)
Call SelectCheckbox("F150V-XSTRF","1",DT_F150_1000_START_IMMEDIATELY_OCC1,True)
Call SelectCheckbox("F150V-XDRAN","1",DT_F150_1000_DUNNPRINT_WITH_SCHEDULING_,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F5\)",True)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

VerifyStatusBar(DT_F150_0100_CHECK_TEXT_OF_STATUSBAR_OCC2)
'Capture the screenshot
Call TakeScreenShot()
Wait(10)
Call GetTextboxValue("F150V-STEXT","3","DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC2_OUTPUT",False)
Call PressEnter() 
wait(2)
Call GetTextboxValue("F150V-STEXT","3","DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC2_OUTPUT",False)
Call GetTextboxValue("F150V-STEXT","4","DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC3_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot()
Call PressEnter() 
wait(2)
Call TakeScreenShot()

Call VerifyTextBoxNoLabelContent("F150V-STEXT","3","    1 dunning notices generated, of which 1 are to be sent",False)
Call VerifyTextBoxNoLabelContent("F150V-STEXT","4","    1 dunning notices printed (cust. 1, vendor 0)",False)

Call ClickButton("Schedule sample dunn\.printout\.\.\.   \(F8\)",False)
Wait(3)

Call SetTextbox("Output Device","USR01-SPLD","",DT_F150_1100_OUTPUT_DEVICE_OCC2,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Start date","F150V-STRDT","",Replace((DT_F150_1000_START_DATE_OCC1),"/","."),True)
Call SelectCheckbox("F150V-XSTRF","1",DT_F150_1000_START_IMMEDIATELY,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Print   \(Ctrl\+P\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

VerifyStatusBar(DT_F150_0100_CHECK_TEXT_OF_STATUSBAR_OCC3)
'Capture the screenshot
Call TakeScreenShot()
Wait(10)
Call PressEnter() 
wait(2)
Call TakeScreenShot()

'Call VerifyTextBoxNoLabelContent("F150V-STEXT","6",DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC5,False)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

