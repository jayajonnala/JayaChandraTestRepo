
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_19-P2P Rebates_ZMU2 - Auto Invoice -R2F-
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0113-Number Range for manual GR done in SAP_P2_cncl_GR
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_0113-Number Range for manual GR done in SAP_P2_cncl_GR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\TASE_DT_P2P_01_01_0113-Number Range for manual GR done in SAP_P2_cncl_GR.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()


Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0, DT_MIGO_2010_GODYNPROMAT_DOC, False)

Call PressEnter()

Call TakeScreenShot()
Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)

Call ClickButton("Check Entries   \(F7\)",False)


Call ClickButton("Post Document   \(Shift\+F11\)",False)

Call TakeScreenShot()

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
'Call VerifyStatusBar("Article document "& DT_GR_NUMBER_OUTPUT &" posted")


Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)

'Call  SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_2010_GODYNPROMAT_DOC_OCC1, False)

Call TakeScreenshot()

Call PressEnter()

Call GetTextboxValue("GOHEAD-LFSNR","","DT_MIGO_0110_CHECK_TEXT_OF_DELIVERY_NOTE_OUTPUT",False)

Call PressEnter()

Call SetTcode(DT_MIGO_0001_OKCD)
Call PressEnter()

Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC2)

Call PressEnter()

Call SetTextBox("Article","GOITEM-MAKTX",0,DT_MIGO_0310_MOVEMENT_TYPE,False)

Call SelectTab("TS_GOITEM", "Quantity",False)
Call SetTextBox("Qty in Unit of Entry","GOITEM-ERFMG",0,DT_MIGO_0315_MOVEMENT_TYPE,False)

Call PressEnter()

Call SelectTab("TS_GOITEM", "Where",False)

Call SetTextBox("Movement type","GOITEM-BWART",0,DT_MIGO_0325_MOVEMENT_TYPE,False)
Call SetTextBox("Site","GOITEM-NAME1",0,DT_MIGO_0325_SITE,False)
Call SetTextBox("Storage Location","GOITEM-LGOBE",0,DT_MIGO_0325_STORAGE_LOCATION,False)

Call PressEnter()

Call TakeScreenShot()
Call SelectTab("TS_GOITEM","Partner",False)

Call TakeScreenShot()
Call SetTextBox("Vendor","GOITEM-VENDORNAME",0,DT_MIGO_0340_VENDOR,False)

Call TakeScreenShot()
Call PressEnter()

Call ClickButton("btn\[11\]",False)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
'Call VerifyStatusBar("Article Document "& DT_GR_NUMBER_2_OUTPUT & " posted")

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC3)
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2010_GODYNPROMAT_DOC_OCC2, False)

Call PressEnter()
Call TakeScreenShot()

Call GetTextboxValue("GOHEAD-MTSNR","","DT_MIGO_0112_CHECK_TEXT_OF_ARTICLE_SLIP_OUTPUT",False)
Call TakeScreenShot()

Call LogOff()

Call FinalStatus()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''










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




