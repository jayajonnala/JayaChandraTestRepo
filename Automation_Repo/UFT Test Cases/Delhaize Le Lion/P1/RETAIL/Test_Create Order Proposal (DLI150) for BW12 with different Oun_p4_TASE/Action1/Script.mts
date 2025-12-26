

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Create Order Proposal (DLI150) for BW12 with different Oun_p4
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

gstrTestCaseName = "Test_Create Order Proposal (DLI150) for BW12 with different Oun_p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''--------TransactionCode-ME23N----------''''
 
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL", "Partners", False)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)
' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)
CALL GetTableCellData("SAPDV70ATC_NAST3","Output Type",1,"","","DT_OutputType1",FALSE)
' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
'VerifyTableCellContent
' GetInputFromExcel(InputFilePath, sheetName, IterationIndex)
call GetInputFromExcel(gstrInputExcelFilePathAndName,"Output",DataRowSet)

Call VerifyTableCellContent(1,"Status", "SAPDV70ATC_NAST3", FormatBlank(DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_0))
Call VerifyTableCellContent(4,"Status", "SAPDV70ATC_NAST3", FormatBlank(DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_3))
Call VerifyTableCellContent(1,"Output Type", "SAPDV70ATC_NAST3", DT_OutputType1)
'Call VerifyTableCellContent(1,"Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(4,"Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)


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


