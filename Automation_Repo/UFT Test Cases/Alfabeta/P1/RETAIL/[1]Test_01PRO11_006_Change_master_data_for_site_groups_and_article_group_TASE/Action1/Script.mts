

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
'.................Test Script Name :Test_01PRO11_006_Change_master_data_for_site_groups_and_article_group
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_01PRO11_006_Change_master_data_for_site_groups_and_article_group"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRO11_006_Change_master_data_for_site_groups_and_article_group_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'''--------------------------------------------  VBG2------------------------------------------------
Call SetTextbox("Article grouping","KONMATGRP-GRPGNR","",DT_VBG2_100_ARTICLE_GROUPING,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call SetTableData("SAPMV23NGRPG_TCTRL","Article",3,"","",DT_VBG2_1100_TABLECELL_ARTICLE_1,False)
Call SetTableData("SAPMV23NGRPG_TCTRL","Unit",3,"","",DT_VBG2_1100_TABLECELL_UOM_1,False)
Call TakeScreenShot()

Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'
'''--------------------------------------------  WB66 .CL24N ------------------------------------------------
Call SetTcode(DT_VBG2_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Copy   \(Enter\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Class","RMCLF-CLASN","",DT_VBG2_1110_CLASS,False)
Call TakeScreenShot()
Call  ClickButton("Change all assignments   \(Ctrl\+F7\)",False)
Call GetTableCellData("SAPLCLFMTC_OBJ_CLASS","Plant / Class",2,"","","DT_CURRENT_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)
Call SelectRowGuiTable("SAPLCLFMTC_OBJ_CLASS","Plant / Class",DT_CURRENT_OUTPUT,False)
Call TakeScreenShot()
Call ClickButton("Delete assignment",False)
Call TakeScreenShot()
Call ClickButton("Yes",True)
'
Call ClickButton("Create new assignments   \(Ctrl\+F9\)",False)
Call SelectRadioButtonIfExist("RMCLF-RADIO",1,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call GetTextboxValue("RMCLF-PAGPOS",0,"ENTRY_OUTPUT",False)
Wait 10

Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)
Wait 10
Call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class",Cint(ENTRY_OUTPUT)+1,"","",DT_NEW,False)
Call PressEnter()     ' 
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

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




