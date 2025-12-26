


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
'.................Test Script Name :Test_04-01-01-01-01-05-Send PO to WMS
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_04-01-01-01-01-05-Send PO to WMS"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-05-Send PO to WMS_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''--------TransactionCode-ME21N----------''''
 
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL", "Partners", False)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)
' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)
'call GetTableCellData("SAPDV70ATC_NAST3","Status",5,"","","DT_Status",False)

'call GetTableCellData("SAPDV70ATC_NAST3","Status",4,"","","DT_Status",False)
'call GetTableCellData("SAPDV70ATC_NAST3","Output Type",4,"","","DT_OutputType",False)
'call GetTableCellData("SAPDV70ATC_NAST3","Medium",4,"","","DT_Medium",False)

' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
'VerifyTableCellContent
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Output",DataRowSet)
Call VerifyTableCellContent(4,"Status", "SAPDV70ATC_NAST3", DT_ME21N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_3)
Call VerifyTableCellContent(4,"Output Type", "SAPDV70ATC_NAST3", DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)
Call VerifyTableCellContent(4,"Medium", "SAPDV70ATC_NAST3", DT_ME21N_0100_CHECK_VALUE_OF_TABLECELL_MEDIUM_3)
Call VerifyTableCellContent(4,"Description", "SAPDV70ATC_NAST3", LCase(DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_3))
Call VerifyTableCellContent(4,"Partner", "SAPDV70ATC_NAST3", DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_PARTNER_3)
Call GetTableCellData("SAPDV70ATC_NAST3","Time",4,"","","DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_TIME_3_OUTPUT",False)
'Call VerifyTableCellContent(5,"Time", "SAPDV70ATC_NAST3", DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_TIME_3_OUTPUT)
Call SelectRowGuiTableByRow("SAPDV70ATC_NAST3",4,False)
Call ClickButton("Processing log   \(Ctrl\+F2\)",False)
Call GetLabelContentByRefLabel("Message text",0,-96,"DT_ME21N_0120_CHECK_TEXT_OF_IDOC_SENT_TO_SAP_SYSTEM_OUTPUT",False)
'Call GetLabelContentByRefLabel("Message text",0,-108,"DT_ME21N_0120_CHECK_TEXT_OF_IDOC_SENT_TO_SAP_SYSTEM_OUTPUT",False)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",FALSE)


''--------TransactionCode-WE02----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
' 
Call SetTcode(DT_ME21N_0100_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_ME21N_0100_OKCD)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_ME21N_1100_CREATED_ON),False)
Call SetTextbox("IDoc Number","DOCNUM-LOW","",DT_ME21N_1100_IDOC_NUMBER,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call VerifyTextBoxContent("Current Status","EDIDC-STATUS",0,DT_ME21N_0100_CHECK_TEXT_OF_CURRENT_STATUS,False)
Call VerifyTextBoxContent("Basic type","EDIDC-IDOCTP",0,DT_ME21N_0100_CHECK_TEXT_OF_BASIC_TYPE,False)
Call VerifyTextBoxContent("Message Type","EDIDC-MESTYP",0,DT_ME21N_0100_CHECK_TEXT_OF_MESSAGE_TYPE,False)
Call VerifyTextBoxContent("Partner No\.","EDPP1-PARNUM",0,DT_ME21N_0100_CHECK_TEXT_OF_PARTNER_NO,False)
Call VerifyTextBoxContent("Extension","EDIDC-CIMTYP",0,DT_ME21N_0100_CHECK_TEXT_OF_EXTENSION,False)


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



