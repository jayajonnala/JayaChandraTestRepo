
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_049-Cancelation Return non-valuated empties
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




gstrTestCaseName = "[1]Test_P2P_01_01_049-Cancelation Return non-valuated empties"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_049-Cancelation Return non-valuated empties.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'''''''''MIGO
Call SetTcode(DT_SAPTRANSACTIONCODE)

Call PressEnter()

Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)

Call TakeScreenShot()

Call ClearTextBox("GODYNPRO-MAT_DOC")

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_2010_GODYNPROMAT_DOC, False)

Call TakeScreenShot()

Call PressEnter()

'Call SelectCellGuiTable("SAPLMIGOTV_GOITEM", "OK", "EUn", "EA", False)

'Call SelectCheckBoxTableByRefColumn("SAPLMIGOTV_GOITEM", "OK", "EUn", "EA", "ON")
' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)


Call SelectCheckBox("GODYNPRO-DETAIL_TAKE",0,"ON",False)

Call ClickButton("Check Entries   \(F7\)", False)


Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" posted")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC1)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT, False)

Call PressEnter()

Call TakeScreenShot()

Call VerifyTableCellContent(1, "Article", "SAPLMIGOTV_GOITEM", DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)

Call VerifyTableCellContent(1, "Movement type", "SAPLMIGOTV_GOITEM", DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)

Call VerifyTableCellContent(1, "Direction", "SAPLMIGOTV_GOITEM", DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)

Call TakeScreenShot()

Call SelectTab("TS_GOHEAD","Doc. info",False)

Call ClickButton("FI Documents",False)

Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC2)

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''














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




