

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Article Creation via DAP_v2_Link empty to ZVOL article
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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

gstrTestCaseName = "Test_Article Creation via DAP_v2_Link empty to ZVOL article"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_002_Create_Purchasing_Conditions_for_an_Existing_Article_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'''----------------------------------
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''----------------------Tcode-MM42 ----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call TakeScreenShot()
Call PressEnter()
wait(2)

Call ClickButton("Goto Component Creation",False)

Call TakeScreenShot()
Call PressEnter()
wait(2)

Call TakeScreenShot()
Call PressEnter()
wait(2)

Call SetTextbox("Unit of Measure","WSTR_DYNP-MEINH","",DT_MM42_0100_UNIT_OF_MEASURE,False)
Call TakeScreenShot()
Call PressEnter()

'Call SetTableData("SAPLWST1TC_COMPONENTS", "Component", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_0,False)
Call SetTableData("SAPLWST1TC_COMPONENTS", "Component qty", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_QTY_0,False)
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickBUtton("Display item details",False)
Wait 2
Call TakeScreenShot()
Call SetTextbox("Storage Type","ZMDAM_MARAEXT-STORAGE_TYPE","",DT_STORAGE_TYPE,False)

'SAPGuiSession("Session").SAPGuiWindow("Change Article.*").SAPGuiTable("SAPLZMDAM_EXT_SCRTC_LTEXT").ClickCell 1,"#1"
'
'Call ClickButton("Create text",False)
'' SetComboByKey(attachedTextOrComboName, keyValue)
'Call SetComboByKey("Create text in","E")
''Call SetCombo("Create text in", "English")
'Call TakeScreenShot()
'Call ClickButton("Enter   \(Enter\)",True)
'CAll SetTextArea("Test")
'Call TakeScreenShot()
'
'Call ClickButton("Create text",False)
''Call SetCombo("Create text in", "Greek")
'Call SetComboByKey("Create text in","G")
'Call TakeScreenShot()
'Call ClickButton("Enter   \(Enter\)",True)
'CAll SetTextArea("Test")
'Call TakeScreenShot()
'
'
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
'Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE_OCC1,False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call TakeScreenShot()
Call PressEnter()
wait(2)

Call ClickButton("Goto Component Creation",False)

Call TakeScreenShot()
Call PressEnter()
wait(2)

Call TakeScreenShot()
Call PressEnter()
wait(2)

Call SetTextbox("Unit of Measure","WSTR_DYNP-MEINH","",DT_MM42_0100_UNIT_OF_MEASURE_OCC1,False)
Call TakeScreenShot()
Call PressEnter()

'Call SetTableData("SAPLWST1TC_COMPONENTS", "Component", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_0_OCC1,False)
Call SetTableData("SAPLWST1TC_COMPONENTS", "Component qty", 1, "", "", DT_MM42_0100_TABLECELL_COMPONENT_QTY_0_OCC1,False)
Call TakeScreenShot()
Call PressEnter()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
'Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)


'''''----------------------Tcode-MM43 ----------------------------
Call SetTcode(DT_MM42_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE_OCC2,False)
Call TakeScreenShot()

Call SelectRowGuiTable("SAPLMGMWTAB_CONT_0100","Screen description","Basic Data",False)
Call TakeScreenShot()
Call PressEnter()
wait(2)

Call ClickButton("Goto Component Creation",False)

Call SetTextbox("Unit of Measure","WSTR_DYNP-MEINH","",DT_MM42_0100_UNIT_OF_MEASURE_OCC2,False)
Call PressEnter()
Call TakeScreenShot()

Call VerifyTextBoxContent("AUTO_EN_TEST.*", "WSTR_DYNP-PACNR", "", DT_MM42_0104_CHECK_TEXT_OF_AUTO_EN_TEST74, False)
Call VerifyTextBoxContent("Unit of Measure","WSTR_DYNP-MEINH" , "", DT_MM42_0100_CHECK_TEXT_OF_UNIT_OF_MEASURE, False)
Call VerifyTableCellContent(1, "Component", "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_0)
Call VerifyTableCellContent(1, "Component qty", "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_QTY_0)



Call SetTextbox("Unit of Measure","WSTR_DYNP-MEINH","",DT_MM42_0100_UNIT_OF_MEASURE_OCC3,False)
Call PressEnter()
Call TakeScreenShot()

Call VerifyTextBoxContent("AUTO_EN_TEST.*", "WSTR_DYNP-PACNR", "", DT_MM42_0104_CHECK_TEXT_OF_AUTO_EN_TEST74_OCC1, False)
Call VerifyTextBoxContent("Unit of Measure","WSTR_DYNP-MEINH" , "", DT_MM42_0100_CHECK_TEXT_OF_UNIT_OF_MEASURE_OCC1, False)
Call VerifyTableCellContent(1, "Component", "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_0_OCC1)
Call VerifyTableCellContent(1, "Component qty", "SAPLWST1TC_COMPONENTS", DT_MM42_0100_CHECK_TEXT_OF_TABLECELL_COMPONENT_QTY_0_OCC1)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

