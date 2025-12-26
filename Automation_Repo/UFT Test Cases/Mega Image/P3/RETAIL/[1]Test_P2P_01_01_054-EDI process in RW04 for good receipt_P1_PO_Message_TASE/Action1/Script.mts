
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test2_MI_Test_P2P_01_01_054-EDI process in RW04 for good receipt_P1_PO_Message     
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



gstrTestCaseName = "Test2_MI_Test_P2P_01_01_054 receipt_P1_PO_Message"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_054-EDI process in RW04 for good receipt_P1_PO_Message.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''//TransactionCode//
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()    
Call TakeScreenShot

'//Write Quantity to Datatable//
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) ' - Line (9)
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) ' - Line (10)
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Expand Header Ctrl\+F2", false)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)

Call TakeScreenShot

'//Write Quantity to Datatable//
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4", false)
Call SelectTab("ITEM_DETAIL","Conditions",False)

Call GetTextboxValue("KOMP-MGAME", "", "DT_ME23N_6201_GET_TEXT_OF_QUANTITY", false)

Call ClickButton("Next item",False)

Call GetTextboxValue("KOMP-MGAME", "", "DT_ME23N_6201_GET_TEXT_OF_QUANTITY_OCC1", false)

'//Verify Messages//
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_1)
Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)

Call VerifyTableCellContent(5, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_4)
Call VerifyTableCellContent(5, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_4)

'//Generate IDoc//

Call SelectRowGuiTable("SAPDV70ATC_NAST3", "Output Type", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0, false)

Call ClickButton("Processing log   \(Ctrl\+F2\)",False)
Call TakeScreenShot
Call GetLabelContentByRefLabel("Message text", 0, -80, "DT_ME23N_0120_CHECK_TEXT_OF_IDOC_ADDED_AND_PASSED_FOR_OUTPUT", True)
Call VerifyifGuiLabelExists(lcase(DT_ME23N_0120_CHECK_TEXT_OF_IDOC_ADDED_AND_PASSED_FOR_OUTPUT))
Call ClickButtonIfExist("Continue   \(Enter\)",True)

'//SetTransactionCode- ME9F//
Call SetTcode(DT_ME23N_0100_OKCD)     
Call PressEnter()     


'''//display Message ZNEU//
Call SetTextbox("Message Type","S_KSCHL-LOW","",DT_ME23N_1000_MESSAGE_TYPE,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectRowGuiGrid("", 0, "Message Language", "EN", false)
Call ClickButton("Display Message   \(Shift\+F8\)",False)
Call TakeScreenShot


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



