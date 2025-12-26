
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_056-EDI process in RW04 for Returns_P1_Messages_GR     
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_P2P_01_01_056-EDI_P1_Messages_GR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_056-EDI process in RW04 for Returns_P1_Chk_Messages.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'--------TransactionCode-ME23N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)

''--------TransactionCode-MB52----------''''

Call SetTcode(DT_ME23N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButtonIfExist("%_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_ME23N_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_ME23N_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButtonIfExist("Check entries   \(Enter\)",True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call SetTextbox("Storage Location","LGORT-LOW","",DT_ME23N_1000_STORAGE_LOCATION,False)
Call SetTextbox("Site","WERKS-LOW","",DT_ME23N_1000_SITE,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

'Call GetLabelContentByRefLabel("Article", -70, -224, "DT_ME23N_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT", False)
'Call GetLabelContentByRefLabel("Article", -245, -224, "DT_ME23N_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT", False)

''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_ME23N_0120_OKCD)     
Call PressEnter()

Call SetComboByKey("GODYNPRO-ACTION", DT_ME23N_0010_GODYNPROACTION)
Call SetComboByKey("GODYNPRO-REFDOC", DT_ME23N_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_ME23N_0003_PUR_ORDER, False)
Call PressEnter()
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME23N_2000_DELIVERY_NUMBER,False)
Call SelectTab("TS_GOITEM","Where",False)
Call VerifyTextBoxContent("Movement type","GOITEM-BWART","",DT_ME23N_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
Call TakeScreenShot
Call PressEnter()
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call TakeScreenShot
Call GetStatusBar("item1","DT_ME23N_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_ME23N_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" posted") 
Call SetComboByKey("GODYNPRO-ACTION", DT_ME23N_0010_GODYNPROACTION_OCC1)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_ME23N_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS_GOITEM","Output",False)
Call TakeScreenShot
Call ClickButton("Display outputs",false)
Call TakeScreenShot
Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_1)
Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)

''--------TransactionCode-MB90----------''''

Call SetTcode(DT_ME23N_0100_OKCD_OCC1)     
Call PressEnter()

Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_ME23N_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT,False)
Call SetTextbox("Article Doc. Year","PM_MJAHR","",DT_ME23N_1000_ARTICLE_DOC_YEAR,False)
'Call SetTextbox("Processing Mode","PM_VERMO","",DT_MIGO_1000_PROCESSING_MODE,False)
'Call SetTextbox("Sort order","PM_NSORT","",DT_MIGO_1000_SORT_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_ME23N_0120_NO_NAME, False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SelectMenuBar("Goto;List Display")
Call VerifyifGuiLabelExists(lcase(DT_ME23N_0120_CHECK_TEXT_OF_RETURN_DELIVERY_OF_GOODS))


Call LogOff()

Call FinalStatus ()







'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



