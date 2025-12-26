'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_07.01.05.01.17 Operational shrinkage"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.01.05.01.17 Operational shrinkage for Goods Receipt Store to_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'---------------------------------MIGO------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenshot()

Call SetTextbox("Movement type","BWART-LOW","",DT_MB51_1000_MOVEMENT_TYPE,False) 
Call SetTextbox("to","BWART-HIGH","",DT_MB51_1000_TO,False) 
Call TakeScreenshot()

Call SetTextbox("Posting Date","BUDAT-LOW","",ConvertDate(DT_MB51_1000_POSTING_DATE),False) 
Call SetTextbox("to","BUDAT-HIGH","",ConvertDate(DT_MB51_1000_TO_OCC1),False)

Call SetTextbox("Site","WERKS-LOW","","",False) 
Call SetTextbox("Article","MATNR-LOW","","",False)
Call TakeScreenshot()

' SetTextboxNoLabel(textboxName, textboxIndex, textboxValue, blnIsItPopup)
''Call SetTextboxNoLabel("BKTXT-LOW","",DT_DOCUMENT_1,False) 
'''Call SetTextboxNoLabel("MBLNR-LOW","",DT_DOCUMENT_1,False) 
Call TakeScreenshot()
Call PressEnter()     
call ClickButton("Execute   \(F8\)",false)
Call TakeScreenshot()


call ClickButtonifexist("Detail List   \(Ctrl\+Shift\+F12\)",false)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
CALL VerifyGridCellContent("",1,"Movement type",0,DT_MOVEMENT_TYPE_1)
CALL VerifyGridCellContent("",1,"Movement type",0,DT_MOVEMENT_TYPE_2)
Call TakeScreenshot()
' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
call GetGridContent("",0,"Article Document",3,"Movement type",DT_MOVEMENT_TYPE_1,"DT_SHRINKAGE_DOCUMENT")

' GetGridContentByRefColumn(gridTitle, gridIndex, refColumn, refFieldVal, columnName, dataTableColumnName)
'GetGridContentByRefColumn("",0,"Reference",
Call TakeScreenshot()

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


