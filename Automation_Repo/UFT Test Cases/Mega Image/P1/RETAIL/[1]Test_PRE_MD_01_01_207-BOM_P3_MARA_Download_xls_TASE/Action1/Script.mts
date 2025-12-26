
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_MD_01_01_207-BOM_P3_MARA_Download_xls
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

gstrTestCaseName = "Test_PRE_MD_01_01_207-BOM_P3_MARA_Download_xls"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_PRE_MD_01_01_207-BOM_P3_MARA_Download_xls.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-SE16----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Table Name","DATABROWSE-TABLENAME","",DT_SE16_0230_TABLE_NAME,false)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("%_I1_%_APP_%-VALU_PUSH", False)
Call TakeScreenShot

Call ClickButtonIfExist("Import from Text File   \(Shift\+F11\)", True)
Call TakeScreenShot

Call SetTextbox("Directory","DY_PATH","",DT_SE16_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_SE16_0200_FILE_NAME,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot

''Call SetTextbox("MTART","I5-LOW","",DT_SE16_1000_ARTICLE_TYPE,false)
Call SetTextboxNoLabel("I5-LOW","",DT_SE16_1000_ARTICLE_TYPE,false)
''Call SetTextbox("MEINS","I10-LOW","",DT_SE16_1000_BASE_UNIT,false)
''Call SetTextbox("MSTAE","I9-LOW","","",false)
'Call SetTextboxNoLabel("I9-LOW","","",false)
Call SetTextbox("Maximum No. of Hits","MAX_SEL","",DT_SE16_1000_MAXIMUM_NO_OF_HITS,false)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)

Call SelectMenuBar("Edit;Download...")  
''this functionality changed, used below button for same functionality
''Call ClickButton("btn\[45\]",False) ''export /download
Call Wait(2)

Call SelectRadioButton("SPOPLI-SELFLAG", "Text with Tabs", True)

Call PressEnter()

Call SetTextbox("Directory","DY_PATH","",DT_SE16_0200_DIRECTORY_OCC1,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_SE16_0200_FILE_NAME_OCC1,True)
Call TakeScreenShot

Call ClickButtonIfExist("Replace existing file   \(Ctrl\+S\)",True)

Call GetStatusBar("messagetype","STATUS_BAR_MESSAGE_TYPE_S_OUTPUT")

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



