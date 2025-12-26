

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_ABI070_001 Create ZCEM Retail Cashier  via ABI070 Interface (
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

gstrTestCaseName = "Test_MD_ABI070_001 Create ZCEM Retail Cashier  via ABI070 Interface ("
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI070_001 Create ZCXT Retail Cashier  via ABI070 Interface (BOS to RETAIL)_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

DtCreatedAt= ConvertDoubledigit(CSTr(Hour(Time)-6))+":" +ConvertDoubledigit(CSTR(Minute(Time)))+ ":00" '--Ccheck the timing
Call SetTextbox("Created At","CRETIM-LOW","",DtCreatedAt,False) 
Call SetTextbox("to","CRETIM-HIGH","",DT_WE05_1100_TO,False)
Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE05_1100_CREATED_ON),False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE05_1100_MESSAGE_VARIANT,False)
Call SetTextboxNoLabel("MESFCT-LOW",0,DT_WE05_1100_MESSAGE_FUNCTION,False)
Call SetTextbox("Partner Port","PPPOR-LOW","",DT_WE05_1100_PARTNER_PORT,False)
Call SetTextbox("Partner Type","PPPRT-LOW","",DT_WE05_1100_PARTNER_TYPE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

'Call VerifyGridCellContent("Selected IDocs",1,"LONG_ICON",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_STATUSICON)
'Call VerifyGridCellContent("Selected IDocs",2,"LONG_ICON",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_STATUSICON)
'Call VerifyGridCellContent("Selected IDocs",3,"LONG_ICON",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_STATUSICON)
'Call VerifyGridCellContent("Selected IDocs",4,"LONG_ICON",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_STATUSICON)
'Call VerifyGridCellContent("Selected IDocs",5,"LONG_ICON",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_STATUSICON)
'Call VerifyGridCellContent("Selected IDocs",6,"LONG_ICON",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_STATUSICON)

Call VerifyGridCellContent("Selected IDocs",1,"IDoc Status",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_IDOCSTATUS)
Call VerifyGridCellContent("Selected IDocs",2,"IDoc Status",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_IDOCSTATUS)
Call VerifyGridCellContent("Selected IDocs",3,"IDoc Status",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_IDOCSTATUS)
Call VerifyGridCellContent("Selected IDocs",4,"IDoc Status",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_IDOCSTATUS)
Call VerifyGridCellContent("Selected IDocs",5,"IDoc Status",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_IDOCSTATUS)
Call VerifyGridCellContent("Selected IDocs",6,"IDoc Status",0,DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCEL_IDOCSTATUS)

Call SelectRowGuiGridbyRowNo("Selected IDocs",0,1,False)
Call ClickButtonToolBar("IDOC",0)
Call TakeScreenShot()
Call ClickLinkGuiTree("#1;#2;#1;#2","#1",0,False)
Call VerifyTableCellContent(46,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE05_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_9)
Call ClickButton("Back   \(F3\)",False)

Call SelectRowGuiGridbyRowNo("Selected IDocs",0,2,False)
Call ClickButtonToolBar("IDOC",0)
Call TakeScreenShot()
Call ClickLinkGuiTree("#1;#2;#1;#2","#1",0,False)
Call VerifyTableCellContent(46,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE05_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_8)
Call ClickButton("Back   \(F3\)",False)

Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



