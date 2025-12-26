

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
'.................Test Script Name :Test_04-01-01-01-01-06-Send PO to DESMAN (DLI030) &C3_V2
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_04-01-01-01-01-06-Send PO to DESMAN (DLI030) &C3_V2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-06-Send PO to DESMAN (DLI030) &C3_V2_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call SetTextboxNoLAbel("I4-LOW","",DT_SE16_1000_WRICEF_ID,false)
Call SetTextboxNoLAbel("I1-LOW","",DT_SE16_1000_PURCHASING_DOC,false)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

'
Call ClickButton("Refresh   \(Ctrl\+F4\)",False)
''' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MANDT,"wnd\[0\]/usr/lbl\[3,5\]")
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN,"wnd\[0\]/usr/lbl\[9,5\]")
call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERDAT),"wnd\[0\]/usr/lbl\[20,5\]")
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRICEF,"wnd\[0\]/usr/lbl\[40,5\]")
call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SENT_IND,"wnd\[0\]/usr/lbl\[51,5\]")

'Call SetTextbox("EBELN.","I1-LOW","",DT_SE16_1000_PURCHASING_DOC,false)
'Call ClickButton("Refresh   \(F8\)",False)
'''
'Call VerifyGridCellContentbyName("shell",2,"MANDT","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MANDT)
'Call VerifyGridCellContentbyName("shell",2,"EBELN","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN)
'Call VerifyGridCellContentbyName("shell",2,"ERDAT","",ConvertDate(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERDAT))
'Call VerifyGridCellContentbyName("shell",2,"WRICEF","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRICEF)
'Call VerifyGridCellContentbyName("shell",2,"SENT_IND","",DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SENT_IND)
'

Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



