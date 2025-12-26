
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AT_P2P_01_01_0261 - Self cons for marketing dept 65-11_P3     
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



gstrTestCaseName = "Test_AT_P2P_01_01_0261  dept 65-11_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="TASE_DT_P2P_01_01_0261-Self consumption for marketing department in DC for consumables SL0004_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''--------TransactionCode-VF01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT", "Document", 1, "", "", DT_VF01_0102_TABLECELL_DOCUMENT_0, false)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1", "DT_VF01_0101_BILLING_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Document "&DT_VF01_0101_BILLING_DOCUMENT_OUTPUT&" has been saved")

'''''--------TransactionCode-/nVF03----------''''

Call SetTcode(DT_VF01_0102_OKCD)     
Call PressEnter()
  
Call SetTextbox("Billing document","VBRK-VBELN","",DT_VF01_0101_BILLING_DOCUMENT_OUTPUT,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call ClickButtonIfExist("Accounting overview   \(Shift\+F4\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)

Call GetTextboxValue("BKPF-BELNR", 0, "DT_VF01_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", False)

Call VerifyGridCellContent("", 1, "BUKRS", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)

Call VerifyGridCellContent("", 1, "BSCHL", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 1, "Account", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("", 1, "PRCTR", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 3, "PRCTR", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)

'''''--------TransactionCode-/nMB90----------''''

Call SetTcode(DT_VF01_0750_OKCD)     
Call PressEnter()
  
Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_VF01_1000_ARTICLE_DOCUMENT,False)
Call SetTextbox("Article Doc. Year","PM_MJAHR","",DT_VF01_1000_ARTICLE_DOC_YEAR,False)
Call SetTextbox("Processing Mode","PM_VERMO","",DT_VF01_1000_PROCESSING_MODE,False)
Call SetTextbox("Sort order","PM_NSORT","",DT_VF01_1000_SORT_ORDER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_VF01_0120_NO_NAME, False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
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


