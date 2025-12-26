
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_059-EDI process in R001_P2_Display_Print_TASE
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_PRE_MI10_Add_Stock_to_SLOC
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)



gstrTestCaseName = "Test_P2P_01_01_059-EDI_P2_Display_Print_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_059-EDI process in R001_P2_Display_Print.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot()
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call PressEnter() 
Call SetTextBoxNoLabel("GODYNPRO-MAT_DOC","0",DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter()
Call SelectTab("TS_GOHEAD","General", False)
Call VerifyTextBoxContent("Delivery Note","GOHEAD-LFSNR",1,DT_MIGO_0110_CHECK_TEXT_OF_DELIVERY_NOTE,False)
Call ClickButtonIfExist("Open Detail Data",false)
Call SelectTab("TS_GOITEM","Where", False)
Call VerifyTextBoxContent("Movement type","GOITEM-BWART",1,DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
Call SelectTab("TS_GOHEAD","Doc. info", False)
Call ClickButtonIfExist("FI Documents",False)

If SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguiedit("guicomponenttype:=31","name:=BKPF-BELNR","attachedtext:=Document Number","Index:=0").Exist(1) Then
	wait 1
Else
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
End If

Call TakeScreenShot


Call VerifyGridCellContent("", 1, "BSCHL", 0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call ClickButton("Back   \(F3\)",False)
If SAPGuisession(sessionObject).sapguiwindow(WindowObject).sapguiedit("guicomponenttype:=32","name:=GOHEAD-CPUDT","attachedtext:=Created On","Index:=0").Exist(1) Then
wait 1
Else
Call ClickButton("Cancel   \(F12\)",True)
End If


Call ClickButtonIfExist("Open Detail Data",false)
Call SelectTab("TS_GOITEM","Output", False)
Call TakeScreenShot
Call ClickButton("Display outputs",False)
Call TakeScreenShot
Call VerifyTableCellContent(1,"Status", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1,"Output Type", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call SetTcode(DT_MIGO_0100_OKCD)
Call PressEnter()
Call TakeScreenShot
Call SetTextBox("Sort order","PM_NSORT",0,DT_MIGO_1000_SORT_ORDER,False)
Call SetTextBox("Processing Mode","PM_VERMO",0,DT_MIGO_1000_PROCESSING_MODE,False)
Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_MIGO_1000_ARTICLE_DOC_YEAR,False)
Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MIGO_1000_ARTICLE_DOCUMENT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckBoxNoLabel(0,"ON",False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Process   \(Shift\+F2\)",False)
Call VerifyStatusBar(DT_MIGO_0120_CHECK_TEXT_OF_STATUSBAR)
Call SetTcode(DT_TRANSACTION_MB5S)
Call PressEnter()
Call TakeScreenShot

Call SetTextboxNolabel("LIFNR-LOW","",DT_ZMDPU_MB5S_1000_VENDOR,False)
Call SetTextbox("Purch\. Organization","EKORG-LOW","",DT_ZMDPU_MB5S_1000_PURCH_ORGANIZATION,False)

Call SetTextbox("Movement type","S_BWART-LOW","",DT_ZMDPU_MB5S_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Reference","S_XBLNR-LOW","",DT_ZMDPU_MB5S_1000_REFERENCE,False)
Call SetTextbox("Posting Date","S_BUDAT-LOW","",ConvertDate(DT_ZMDPU_MB5S_1000_POSTING_DATE),False)
Call SetTextbox("Entry Date","S_EDATE-LOW","",ConvertDate(DT_ZMDPU_MB5S_1000_POSTING_DATE),False)
Call SetTextbox("Company Code","S_BUKRS-LOW","",DT_COMPANY_CODE,False)
Call SetTextbox("Site","S_WERKS-LOW","",DT_SITE_FROM,False)
Call SetTextbox("to","S_WERKS-HIGH","",DT_SITE_TO,False)
Call ClickButton("Execute   \(F8\)",False)
Call SelectCheckBoxNoLabel(0,"ON",False)
Call SelectCheckBoxNoLabel(1,"ON",False)
Call ClickButton("Send notification   \(Shift\+F2\)",false)
Call VerifyifGuiLabelExists(DT_ZMDPU_MB5S_0120_CHECK_TEXT_OF_NO_NAME)

Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''










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




