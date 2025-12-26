
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0251-Scrapping on a different CC    
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



gstrTestCaseName = "Test_P2P_01_01_0251-Scrapping on a different CC"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0251-Scrapping on a different CC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call SetComboByKey("GODYNPRO-REFDOC", "R10")
Call PressEnter()
'Call SetTextbox("GI scrapping","GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call PressEnter()
Call SelectTab("TS_GOITEM","Article",False)
Call TakeScreenShot
Call SetTextbox("Article","GOITEM-MAKTX","",DT_MIGO_0310_ARTICLE,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS_GOITEM","Quantity",False)
Call TakeScreenShot
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_QTY_IN_UNIT_OF_ENTRY,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TS_GOITEM","Where",False)
Call TakeScreenShot
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_0325_MOVEMENT_TYPE,False)
Call SetTextbox("Site","GOITEM-NAME1","",DT_MIGO_0325_SITE,False)
Call SetTextbox("Storage Location","GOITEM-LGOBE","",DT_MIGO_0325_STORAGE_LOCATION,False)
Call SetTextbox("Reason for Movement","GOITEM-GRUND","",DT_MIGO_0325_REASON_FOR_MOVEMENT,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS_GOITEM","Account Assignment",False)
Call TakeScreenShot
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_MIGO_1001_COST_CENTER,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Check Entries   \(F7\)",false)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")

Call ClickButton("Restart   \(F5\)",false)
Call PressEnter() 
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS_GOITEM","Output",False)
Call TakeScreenShot
Call ClickButton("Display outputs",false)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1, "Description", "SAPDV70ATC_NAST3", lcase(DT_MIGO_0100_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_0))
Call VerifyTableCellContent(1, "Medium", "SAPDV70ATC_NAST3", lcase(DT_MIGO_0100_CHECK_VALUE_OF_TABLECELL_MEDIUM_0))
Call ClickButton("Back   \(F3\)",false)
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",false)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButton("Display Document   \(F2\)",True)
Call TakeScreenShot
Call GetTextboxValue("BKPF-BELNR", 0, "DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", False)
Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

''''''--------TransactionCode-MB90----------''''

Call SetTcode(DT_MIGO_0750_OKCD)     
Call PressEnter()     

Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call SetTextbox("Article Doc. Year","PM_MJAHR","",Year(Date),False)
''Call SetTextbox("Processing Mode","PM_VERMO","",DT_MIGO_1000_PROCESSING_MODE,False)  '''not required
''Call SetTextbox("Sort order","PM_NSORT","",DT_MIGO_1000_SORT_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_MIGO_0120_NO_NAME, False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
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


