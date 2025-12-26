


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
'.................Test Script Name :Test_01PRO11_001_Create_Article_groupings_using_various_select_option
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_01PRO11_001_Create_Article_groupings_using_various_select_option"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRO11_001_Create_Article_groupings_using_various_select_option_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'
''--------------------------------------------  WAK2------------------------------------------------
'
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","","",True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program .*","","Variant name","AB_TEST10",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)
Call ClickButton("Execute   \(F8\)",False) 
Wait 5
Call TakeScreenShot()
Call verifyNoRowExistsGrid("shell",0)
'Call VerifyWebElement(0,"Mass maintenance of Articles in an Article Grouping","SPAN","0","","",False)
Call  ClickButton("Back   \(F3\)",False)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUM",Cint(DT_NUM)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Article grouping","P_GRPGNR","",DT_ZMDPR_ARTICLE_GROUP_1000_ARTICLE_GROUPING,False)
Call SetTextbox("Text article groupng","P_GRPTXT","",DT_ZMDPR_ARTICLE_GROUP_1000_TEXT_ARTICLE_GROUPNG,False)
Call SetTextbox("Article","S_MATNR-LOW","","",False)
'Call SetTextbox("Vendor","S_LIFNR-LOW","",DT_ZMDPR_ARTICLE_GROUP_1000_VENDOR,False)
Call SetTextboxNoLabel("S_LIFNR-LOW","",DT_ZMDPR_ARTICLE_GROUP_1000_VENDOR,False)
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_ZMDPR_ARTICLE_GROUP_1000_PURCHASING_GROUP,False)
Call SetTextbox("Merchandise Category","S_MATKL-LOW","","",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyGridCellContent("Select Items to be Insterted into Article Group",1,"MATNR%",0,DT_ZMDPR_ARTICLE_GROUP_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call  SelectMenuBar("Edit;Select All")
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call TakeScreenShot()

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




