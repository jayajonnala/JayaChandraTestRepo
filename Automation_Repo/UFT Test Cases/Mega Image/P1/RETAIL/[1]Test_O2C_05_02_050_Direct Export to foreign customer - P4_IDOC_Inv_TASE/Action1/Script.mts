
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_05_02_050_Direct Export to foreign customer - P4_IDOC_Inv
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

gstrTestCaseName = "Test_O2C_05_02_050_Direct Export to foreign customer - P4_IDOC_Inv"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_05_02_050_Direct Export to foreign customer - P4_IDOC_Inv.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''--------TransactionCode-WE05----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_DATE),False)
''Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_DATE),False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE02_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE02_1100_MESSAGE_FUNCTION,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_WE05_1300_TRANSFER_FILE_REFERENCE,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE05_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)
Call SelectMenuBar("Goto;Display Links")
Call VerifyGridCellContent("Relationships to IDoc:.*", 1, "Description (Partner)", 0, DT_WE05_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PARTNER)
Call ClickButton("Cancel   \(F12\)",True)

Call SetTcode(DT_WE05_0100_OKCD)
Call PressEnter() 

Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE05_1100_MESSAGE_VARIANT_OCC1,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE05_1100_MESSAGE_FUNCTION_OCC1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
Call TakeScreenShot
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_WE05_1300_TRANSFER_FILE_REFERENCE_OCC2,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE05_0100_CHECK_TEXT_OF_CURRENT_STATUS_OCC1, False)
Call SelectMenuBar("Goto;Display Links")
Call VerifyGridCellContent("Relationships to IDoc:.*", 1, "Description (Partner)", 0, DT_WE05_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PARTNER_OCC1)
Call ClickButton("Cancel   \(F12\)",True)

'''''''--------TransactionCode-/nVL02N----------''''''Additional step''not available in current log''VL02N
Call SetTcode("/nVL02N")     
Call PressEnter()     
Call TakeScreenShot

''ENTER OUTBOUND DELIVERY
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_WE05_0102_TABLECELL_DOCUMENT_0,False)
Call PressEnter()     
Call TakeScreenShot

''Click on Post Goods Issue  button
Call ClickButtonIfExist("Post Goods Issue   \(Shift\+F8\)",False)
Call TakeScreenShot


'''''--------TransactionCode-VF01----------''''

Call SetTcode(DT_WE05_0100_OKCD_OCC1)     
Call PressEnter()     


Call SetTableData("SAPMV60ATCTRL_ERF_FAKT", "Document", 2, "", "", DT_WE05_0102_TABLECELL_DOCUMENT_0, false)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
'Call VerifyStatusBar(DT_WE05_0102_CHECK_TEXT_OF_STATUSBAR)

'DT_WE05_0102_CHECK_TEXT_OF_STATUSBAR
Call GetStatusBar("item1", "DT_WE05_0102_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_WE05_0102_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" has been saved")
Call TakeScreenShot

'''''--------TransactionCode-VF03----------''''

Call SetTcode(DT_WE05_0102_OKCD)     
Call PressEnter()     

Call SetTextbox("Billing document", "VBRK-VBELN", "", DT_WE05_0102_CHECK_TEXT_OF_STATUSBAR_OUTPUT, False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call ClickButtonIfExist("Accounting overview   \(Shift\+F4\)",False)
Call TakeScreenShot
'Updated row number as 1
'Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Doc. Number", True)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Posting Key", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "Account", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 2, "Profit Center", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

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



