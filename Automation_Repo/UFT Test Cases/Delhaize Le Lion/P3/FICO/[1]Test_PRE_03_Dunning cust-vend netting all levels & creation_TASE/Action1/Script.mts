'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_03_Dunning cust-vend netting all levels&creation 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_03_Dunning cust-vend net&creati"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Call CloseSessionsSAP()

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode--F150---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Run On","F150V-LAUFD","",ConvertDate(DT_F150_0100_RUN_ON),False)
Call SetTextbox("Identification","F150V-LAUFI","",DT_F150_0100_IDENTIFICATION,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("DU_TABSTRIP", "Parameter", False)
Call SetTextbox("Docmnts posted up to","F150V-GRDAT","",ConvertDate(DT_F150_0111_DOCMNTS_POSTED_UP_TO),False)
Call SetTextbox("Dunning date","F150V-AUSDT","",ConvertDate(DT_F150_0111_DUNNING_DATE),False)
Call TakeScreenShot
Call SetTextbox("Company Code","RNG_BKRS-LOW","",DT_F150_0001_COMPANY_CODE,False)
Call SetTextbox("Customer","RNG_SELC-LOW","",DT_F150_0002_CUSTOMER,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("DU_TABSTRIP", "Free selection", False)
Call SetTableData("SAPF150VDU_FLDTAB", "Exclude values", 1, "", "","ON", False)
Call SetTableData("SAPF150VDU_FLDTAB", "Field name", 1, "", "",DT_F150_0112_TABLECELL_FIELD_NAME_0, False)
Call SetTableData("SAPF150VDU_FLDTAB", "Field name", 2, "", "",DT_F150_0112_TABLECELL_FIELD_NAME_1, False)
Call SetTableData("SAPF150VDU_FLDTAB", "Field name", 3, "", "",DT_F150_0112_TABLECELL_FIELD_NAME_2, False)
Call SetTableData("SAPF150VDU_FLDTAB", "Values", 1, "", "",DT_F150_0112_TABLECELL_VALUES_0, False)
Call SetTableData("SAPF150VDU_FLDTAB", "Values", 2, "", "",DT_F150_0112_TABLECELL_VALUES_1, False)
Call SetTableData("SAPF150VDU_FLDTAB", "Values", 3, "", "",DT_F150_0112_TABLECELL_VALUES_2, False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)

Call SelectTab("DU_TABSTRIP", "Status", False)
Call ClickButtonIfExist("Yes", True)
Call ClickButtonIfExist("Schedule dunning run   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("Output Device","USR01-SPLD","",DT_F150_1100_OUTPUT_DEVICE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SelectCheckbox("F150V-XSTRF", 0,DT_F150_1000_START_IMMEDIATELY, True)
Call TakeScreenShot
'''Call ClickButtonIfExist("Execute   \(F5\)", True)

Call ClickButton("Execute   \(F5\)",True)
Call PressEnter()  
Call PressEnter()  
Call PressEnter()  
Call TakeScreenShot

'''Call ClickButtonIfExist("Dunning list   \(Shift\+F6\)",False)
Call ClickButton("Dunning list   \(Shift\+F6\)",False)
Call SetTextbox("Program","F150V-PROGN","",DT_F150_1113_PROGRAM,True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_F150_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[69,10\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[1,11\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[1,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC3,"wnd\[0\]/usr/lbl\[17,11\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC4,"wnd\[0\]/usr/lbl\[17,12\]")

'''Call VerifyifGuiLabelExists_ByIndex(DT_F150_0120_CHECK_TEXT_OF_NO_NAME, 0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC1, 0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC2, 0)
''Call VerifyifGuiLabelExists_ByIndex(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC3, 0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC4, 0)

Call ClickButtonIfExist("Back   \(F3\)",False)

'''Call ClickButtonIfExist("Schedule sample dunn\.printout\.\.\.   \(F8\)",False)
Call ClickButton("Schedule sample dunn\.printout\.\.\.   \(F8\)",False)
Call SetTextbox("Output Device","USR01-SPLD","",DT_F150_1100_OUTPUT_DEVICE_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SelectCheckbox("F150V-XSTRF", 0,DT_F150_1000_START_IMMEDIATELY_OCC1, True)
Call TakeScreenShot
Call ClickButtonIfExist("Display   \(F7\)", True)
Call TakeScreenShot

Call SelectMenuBar("Goto;List Display")
Call TakeScreenShot
Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
Call SetTextbox("Find","RSYSF-STRING","",DT_F150_0120_CHECK_TEXT_OF_VAN_EUR__2000_00_VERSCHULDIGD,True)
Call ClickButtonIfExist("Find   \(Enter\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex("1", 0)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call SetTextbox("Find","RSYSF-STRING","",DT_F150_0120_CHECK_TEXT_OF_REFERENTIE_D1_11005829__EN_DIT,True)
Call ClickButtonIfExist("Find   \(Enter\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex("1", 0)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call SetTextbox("Find","RSYSF-STRING","",DT_F150_0120_CHECK_TEXT_OF_BINNEN_DE_8_DAGEN_NA,True)
Call ClickButtonIfExist("Find   \(Enter\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex("1", 0)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call SetTextbox("Find","RSYSF-STRING","",DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC5,True)
Call ClickButtonIfExist("Find   \(Enter\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex("1", 0)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call SetTextbox("Find","RSYSF-STRING","",DT_F150_0120_CHECK_TEXT_OF_NO_NAME_OCC7,True)
Call ClickButtonIfExist("Find   \(Enter\)", True)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex("1", 0)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call ClickButtonIfExist("Cancel   \(F12\)", True)
Call ClickButtonIfExist("Next Page   \(Page down\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Next Page   \(Page down\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)

''Call ClickButtonIfExist("Schedule sample dunn\.printout\.\.\.   \(F8\)",False)
Call ClickButtonIfExist("Schedule dunn\.notice printout   \(F6\)",False)

Call SetTextbox("Output Device","USR01-SPLD","",DT_F150_1100_OUTPUT_DEVICE_OCC1,True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SelectCheckbox("F150V-XSTRF", 0,DT_F150_1000_START_IMMEDIATELY_OCC1, True)
Call TakeScreenShot
Call ClickButtonIfExist("Print   \(Ctrl\+P\)", True)
Call TakeScreenShot

Call GetStatusBar("item1", "DT_OP_F150_0110_JOB_TEXT_OF_F150VSJOB_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_OP_F150_0110_JOB_TEXT_OF_F150VSJOB_OUTPUT",DT_OP_F150_0110_JOB_TEXT_OF_F150VSJOB)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT)
Call PressEnter()
Call TakeScreenShot

Call VerifyTextBoxNoLabelContent("F150V-STEXT",2,DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC1,False)
Call VerifyTextBoxNoLabelContent("F150V-STEXT",5,DT_F150_0110_CHECK_TEXT_OF_F150VSTEXT_OCC2,False)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

