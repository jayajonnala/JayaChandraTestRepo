'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S05V18 AR02 Asset History Sheet 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S05V18 AR02 As His"
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
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-AR03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company code","BUKRS-LOW","",DT_AR03_1000_COMPANY_CODE,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_AR03_1000_SORT_VARIANT,False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_AR03_1000_DEPRECIATION_AREA,False)
Call SetTextbox("Report date","BERDATUM","",ConvertDate(DT_AR03_1000_REPORT_DATE),False)
Call SetTextbox("Asset class","SO_ANLKL-LOW","","",False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_AR03_1000_ASSET_NUMBER,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyifGuiLabelExistsByRelativeid("004/000","wnd\[0\]/usr/lbl\[36,11\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_AR03_0120_CHECK_TEXT_OF_NO_NAME_OCC5,"wnd\[0\]/usr/lbl\[87,11\]")

Call GetLabelContentByRefLabel("CompanyCode","-728","-128","DT_OP_AR03_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

'''''--------TransactionCode-ZFIAA_BUK_DEPEND----------''''

Call SetTcode(DT_AR03_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Position\.\.\.",False)
Call SetTextbox("Company Code","SVALD-VALUE","",DT_AR03_0300_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Display \-\> Change   \(Ctrl\+F1\)",False)
Call ClickButton("Position\.\.\.",False)
Call SetTextbox("Company Code","SVALD-VALUE","",DT_AR03_0300_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call SetTableData("SAPLZFIAA_TAB_MAINTTCTRL_ZFIAA_V_BUK_DEPE", "Depr. End", 1, "", "",DT_AR03_1000_REPORT_DATE, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
CAll VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_AR03_0100_CHECK_TEXT_OF_STATUSBAR)
Call VerifyTableCellContent(1, "Depr. End","SAPLZFIAA_TAB_MAINTTCTRL_ZFIAA_V_BUK_DEPE",ConvertDate(DT_AR03_0100_CHECK_TEXT_OF_TABLECELL_DEPR_END_0))
Call ClickButton("Back   \(F3\)",False)
'
''''''--------TransactionCode-AR03----------''''

Call SetTcode(DT_AR03_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'

Call SetTextbox("Company code","BUKRS-LOW","",DT_AR03_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_AR03_1000_SORT_VARIANT_OCC1,False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_AR03_1000_DEPRECIATION_AREA_OCC1,False)
Call SetTextbox("Report date","BERDATUM","",ConvertDate(DT_AR03_1000_REPORT_DATE_OCC1),False)
Call SetTextbox("Asset class","SO_ANLKL-LOW","",DT_AR03_1000_ASSET_CLASS_OCC1,False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_AR03_1000_ASSET_NUMBER_OCC1,False)
Call SetTextbox("Simulation version","PA_SIMVR","",DT_AR03_1000_SIMULATION_VERSION,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call VerifyifGuiLabelExistsByRelativeid(DT_AR03_0120_CHECK_TEXT_OF_NO_NAME_OCC5,"wnd\[0\]/usr/lbl\[87,12\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_AR03_0120_CHECK_TEXT_OF_NO_NAME_OCC6,"wnd\[0\]/usr/lbl\[104,12\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_AR03_0120_CHECK_TEXT_OF_NO_NAME_OCC7,"wnd\[0\]/usr/lbl\[121,13\]")
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
