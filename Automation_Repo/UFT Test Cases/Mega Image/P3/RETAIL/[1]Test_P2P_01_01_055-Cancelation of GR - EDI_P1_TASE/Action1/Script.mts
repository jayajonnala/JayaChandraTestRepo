
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_055-Cancelation of GR - EDI_P1     
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

gstrTestCaseName = "Test_P2P_01_01_055-Cancelation of GR - EDI_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_055-Cancelation of GR - EDI_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call TakeScreenShot
Call SetComboByKey("GODYNPRO-REFDOC", DT_MIGO_0010_GODYNPROREFDOC)
Call TakeScreenShot
Call SelectTab("TS_GOITEM","Article",False)
Call TakeScreenShot
Call SetTextbox("Article","GOITEM-MAKTX","",DT_MIGO_0310_MOVEMENT_TYPE,False)
Call SelectTab("TS_GOITEM","Quantity",False)
Call TakeScreenShot
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_MOVEMENT_TYPE,False)
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFME","",DT_MIGO_0315_MOVEMENT_TYPE_UNIT,False)
Call PressEnter()
Call SelectTab("TS_GOITEM","Where",False)
Call TakeScreenShot
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_0325_MOVEMENT_TYPE,False)
Call SetTextbox("Site","GOITEM-NAME1","",DT_MIGO_0325_SITE,False)
Call SetTextbox("Storage Location","GOITEM-LGOBE","",DT_MIGO_0325_STORAGE_LOCATION,False)
''''''next 3-4steps are wrong in log, replacing Movement Type text 3 times
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TS_GOITEM","Partner",False)
Call TakeScreenShot
Call SetTextbox("Vendor","GOITEM-VENDORNAME","",DT_MIGO_0340_VENDOR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("OK_NEXT_ITEM",false)
Call SelectTab("TS_GOITEM","Article",False)
Call TakeScreenShot
Call SetTextbox("Article","GOITEM-MAKTX","",DT_MIGO_0310_MOVEMENT_TYPE_OCC1,False)
Call SelectTab("TS_GOITEM","Quantity",False)
Call TakeScreenShot
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_MOVEMENT_TYPE_OCC1,False)
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFME","",DT_MIGO_0315_MOVEMENT_TYPE_UNIT,False)
Call PressEnter()
Call SelectTab("TS_GOITEM","Where",False)
Call TakeScreenShot
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_0325_MOVEMENT_TYPE,False)
Call SetTextbox("Site","GOITEM-NAME1","",DT_MIGO_0325_SITE,False)
Call SetTextbox("Storage Location","GOITEM-LGOBE","",DT_MIGO_0325_STORAGE_LOCATION,False)

Call TakeScreenShot
Call PressEnter()
Call SelectTab("TS_GOITEM","Partner",False)
Call TakeScreenShot
Call SetTextbox("Vendor","GOITEM-VENDORNAME","",DT_MIGO_0340_VENDOR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Check Entries   \(F7\)",false)
Call VerifyStatusBar(lcase(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR))

Call ClickButton("Post Document   \(Shift\+F11\)",false)

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")


Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)
Call TakeScreenShot

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
''Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR", 0, DT_MIGO_2010_GODYNPRODOC_YEAR, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SelectTab("TS_GOITEM","Output",False)
Call TakeScreenShot
Call ClickButton("Display outputs",false)
Call TakeScreenShot
Call VerifyTableCellContent(3, "Status", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_1)
Call VerifyTableCellContent(3, "Output Type", "SAPDV70ATC_NAST3", DT_MIGO_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)

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



