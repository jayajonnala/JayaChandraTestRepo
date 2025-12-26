

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
'.................Test Script Name :Test_04-01-01-01-01-06-Send PO to DESMAN (DLI030) &C3
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_04-01-01-01-01-06-Send PO to DESMAN (DLI030) &C3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-06-Send PO to DESMAN (DLI030) &C3_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'//--------- SE16----------------------------------------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
 

Call SetTextbox("Table Name","DATABROWSE-TABLENAME","",DT_SE16_0230_TABLE_NAME,false)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextboxNoLabel("I4-LOW","",DT_SE16_1000_WRICEF_ID,false)
Call SetTextboxNoLabel("I1-LOW","",DT_SE16_1000_PURCHASING_DOC,false)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'Call SetTextbox("WRICEF ID","I4-LOW","",DT_SE16_1000_WRICEF_ID,false)
'Call SetTextbox("Purchasing Doc\.","I1-LOW","",DT_SE16_1000_PURCHASING_DOC,false)
Call ClickButton("Refresh   \(Ctrl\+F4\)",False)
'' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MANDT,"wnd\[0\]/usr/lbl\[3,5\]")
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN,"wnd\[0\]/usr/lbl\[9,5\]")
call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERDAT),"wnd\[0\]/usr/lbl\[20,5\]")
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRICEF,"wnd\[0\]/usr/lbl\[40,5\]")
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SENT_IND,"wnd\[0\]/usr/lbl\[51,5\]")
'Call VerifyGridCellContentbyName("shell",1,"EBELN","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN)ERDAT
'

'Call ClickButton("Refresh   \(F8\)",False)
'' VerifyifGuiLabelExists(Content)
'call VerifyifGuiLabelExists(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MANDT)
''' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
''VerifyifGuiLabelExistsByRelativeid
''' VerifyifGuiLabelExists_ByIndex(Content, Index)
''VerifyifGuiLabelExists_ByIndex
'
'Call VerifyGridCellContentbyName("shell",1,"MANDT","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MANDT)
'Call VerifyGridCellContentbyName("shell",1,"EBELN","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN)
'Call VerifyGridCellContentbyName("shell",1,"ERDAT","",ConvertDate(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERDAT))
'Call VerifyGridCellContentbyName("shell",1,"WRICEF","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRICEF)
'Call VerifyGridCellContentbyName("shell",1,"SENT_IND","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SENT_IND)


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



