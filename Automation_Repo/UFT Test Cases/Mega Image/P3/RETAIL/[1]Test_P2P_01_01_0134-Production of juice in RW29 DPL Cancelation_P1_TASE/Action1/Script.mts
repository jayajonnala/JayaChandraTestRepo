
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0134-Production of juice in RW29 DPL Cancelation_P1     
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


gstrTestCaseName = "Test_P2P_01_01_0134- RW29 DPL Cancelation_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0134-Production of juice in RW29 DPL Cancelation.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''--------TransactionCode-mfbf----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Conf. Qty","RM61B-ERFMG","",DT_MFBF_0811_CONF_QTY,False)
Call TakeScreenShot
Call SetTextbox("Posting Date","RM61B-BUDAT","",ConvertDate(DT_MFBF_0800_POSTING_DATE),False)
Call SetTextbox("Document Date","RM61B-BLDAT","",ConvertDate(DT_MFBF_0800_DOCUMENT_DATE),False)
Call TakeScreenShot
Call SetTextbox("Article","RM61B-MATNR","",DT_MFBF_0801_ARTICLE,False)
Call SetTextbox("Site","RM61B-WERKS","",DT_MFBF_0801_SITE,False)
Call SetTextbox("Production Version","RM61B-VERID","",DT_MFBF_0801_PRODUCTION_VERSION,False)
Call PressEnter()
Call TakeScreenShot
'Call ClickButton("btn\[17\]",False)  ''Doc.-Neutral Reversal
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("Text", "DT_MFBF_0800_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_MFBF_0800_CHECK_TEXT_OF_STATUSBAR_OUTPUT)

'''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_MFBF_0800_OKCD)     
Call PressEnter()     

Call SetComboByKey("GODYNPRO-ACTION", DT_MFBF_0010_GODYNPROACTION)
Call PressEnter() 
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MFBF_2010_GODYNPROMAT_DOC, False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("BUTTON_ITEMDETAIL",false)
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call ClickButtonIfExist("BUTTON_ITEMDETAIL",false)
Call VerifyTableCellContent(1, "Movement type", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call VerifyTableCellContent(1, "Direction", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)
Call VerifyTableCellContent(2, "Movement type", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
Call VerifyTableCellContent(2, "Direction", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_1)
Call VerifyTableCellContent(3, "Movement type", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
Call VerifyTableCellContent(3, "Direction", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_2)
Call VerifyTableCellContent(4, "Movement type", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_3)
Call VerifyTableCellContent(4, "Direction", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_3)
Call VerifyTableCellContent(5, "Movement type", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_4)
Call VerifyTableCellContent(5, "Direction", "SAPLMIGOTV_GOITEM", DT_MFBF_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_4)
Call ClickButtonIfExist("BUTTON_DETAIL",false)
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call ClickButton("FI Documents",false)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)
Call TakeScreenShot
Call GetTextboxValue("BKPF-BELNR", "", "DT_MFBF_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", False)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
''Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
''Call VerifyGridCellContent("", 5, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
''Call VerifyGridCellContent("", 6, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 5, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 6, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("", 7, "BSCHL", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
''Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
''Call VerifyGridCellContent("", 4, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
''Call VerifyGridCellContent("", 5, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
''Call VerifyGridCellContent("", 6, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 5, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 6, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("", 7, "KTONR", 0, DT_MFBF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",false)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call TakeScreenShot
Call SelectTab("TS_GOITEM", "Output", False)
Call TakeScreenShot
Call ClickButton("Display outputs",false)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_MFBF_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_MFBF_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

''''--------TransactionCode-/nmb90----------''''

Call SetTcode(DT_MFBF_0100_OKCD)     
Call PressEnter()     

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_MFBF_1000_ARTICLE_DOCUMENT,False)
Call SetTextbox("Article Doc. Year","PM_MJAHR","",DT_MFBF_1000_ARTICLE_DOC_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_MFBF_0120_NO_NAME, False)
Call ClickButton("Print preview   \(Shift\+F4\)",false)
Call TakeScreenShot
Call SelectMenuBar("Goto;List Display")
Call TakeScreenShot

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



