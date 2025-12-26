
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_Get_Article_Doc_PO_History_Filterby_TypeTest_O2C_03_01_018-Distribution process to MI stores - RW04 Dry_Popesti_P3_Chk_Postings      
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


gstrTestCaseName = "Test_O2C_03_01_018-Distribution process to MI stores - RW04 Dry_Popesti_P3_Chk_Postings"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_03_01_018-Distribution process to MI stores-RW04 Dry_Popesti_P3_Chk_Postings.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''--------TransactionCode-WE05----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_DATE),False)
Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_DATE),False)
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
Call ClickButton("Back   \(F3\)",False)
Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
Call TakeScreenShot
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_WE05_1300_TRANSFER_FILE_REFERENCE_OCC1,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE05_0100_CHECK_TEXT_OF_CURRENT_STATUS_OCC1, False)
Call SelectMenuBar("Goto;Display Links")
Call VerifyGridCellContent("Relationships to IDoc:.*", 1, "Description (Partner)", 0, DT_WE05_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PARTNER_OCC1)
Call ClickButton("Cancel   \(F12\)",True)

''--------TransactionCode-VL03N----------''''

Call SetTcode(DT_WE05_0101_OKCD)     
Call PressEnter()     
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_WE05_4004_OUTBOUND_DELIVERY,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Document Flow   \(F7\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#2")
Call TakeScreenShot
Call GetGridContent("TF to stck.*", 0, "DOCNUM", 1, "<NA>", "<NA>", "DT_VL03N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyNodeTextGuiTree(0, DT_WE05_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC1)


'''''--------TransactionCode-MB03----------''''

Call SetTcode(DT_WE05_0100_OKCD)     
Call PressEnter()     
Call SetTextbox("Article Doc.", "RM07M-MBLNR", "", DT_WE05_0460_ARTICLE_DOC, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
wait 5
Call PressEnter()

Call TakeScreenShot
Call ClickButtonIfExist("Accounting Documents   \(F7\)",False)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)

Call VerifyGridCellContent("", 1, "BSCHL", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 1, "KOBEZ", "", lcase(DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ))
Call VerifyGridCellContent("", 2, "KOBEZ", "", lcase(DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ))
Call GetGridContent("", 0, "AZBET", 1, "<NA>", "<NA>", "DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OUTPUT")
Call GetGridContent("", 0, "AZBET", 2, "<NA>", "<NA>", "DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OUTPUT")
Call VerifyGridCellContent("", 1, "Currency", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 1, "GSBER", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call VerifyGridCellContent("", 2, "GSBER", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)
Call VerifyGridCellContent("", 1, "EBELN", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN)
Call VerifyGridCellContent("", 2, "EBELN", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_EBELN)
Call VerifyGridCellContent("", 1, "AUGBL", "", "")
Call VerifyGridCellContent("", 2, "AUGBL", "", "")
Call VerifyGridCellContent("", 1, "VBUND", "", "")
Call VerifyGridCellContent("", 2, "VBUND", "", "")
Call VerifyGridCellContent("", 1, "PRCTR", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

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



