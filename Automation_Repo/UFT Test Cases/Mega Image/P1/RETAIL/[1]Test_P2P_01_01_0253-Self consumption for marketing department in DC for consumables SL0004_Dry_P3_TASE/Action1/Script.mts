
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P1     

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


gstrTestCaseName = "Test_P2P_01_01_0253_Dry_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextBox("Sort order","PM_NSORT",0,DT_MB90_1000_SORT_ORDER,False)
Call SetTextBox("Processing Mode","PM_VERMO",0,DT_MB90_1000_PROCESSING_MODE,False)
'''''''''''''''Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_YEAR,False)
Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MB90_1000_ARTICLE_DOCUMENT,False)

Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Execute   \(F8\)",False)
Call SelectCheckBoxNoLabel(0,"ON",False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MB90_1000_ARTICLE_DOCUMENT_OCC1,False)

Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Execute   \(F8\)",False)
Call SelectCheckBoxNoLabel(0,"ON",False)
Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call SetTcode(DT_MB90_0100_OKCD)
Call PressEnter()
Call SetTableData("SAPMV60ATCTRL_ERF_FAKT", "Document", 1, "", "", DT_MB90_0102_TABLECELL_DOCUMENT_0, False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("btn\[11\]",False)
Call GetStatusBar("item1","DT_MB90_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_MB90_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been saved")
''''''


Call SetTableData("SAPMV60ATCTRL_ERF_FAKT", "Document", 1, "", "", DT_MB90_0102_TABLECELL_DOCUMENT_0_OCC1, False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("btn\[11\]",False)
Call GetStatusBar("item1","DT_MB90_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_MB90_0102_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" has been saved")

Call SetTcode(DT_MB90_0100_OKCD_OCC1)
Call PressEnter()
Call SetTextBox("Article Doc\.","RM07M-MBLNR",0,DT_MB90_0460_ARTICLE_DOC,False)
Call PressEnter()
Call TakeScreenShot()
wait 5
Call PressEnter()
Call ClickButton("Accounting Documents   \(F7\)",False)

Call VerifyGridCellContent("", 1, "Company code", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call SetTextBox("Article Doc\.","RM07M-MBLNR",0,DT_MB90_0460_ARTICLE_DOC_OCC1,False)
Call PressEnter()
Call TakeScreenShot()
wait 5
Call PressEnter()
Call ClickButton("Accounting Documents   \(F7\)",False)

Call VerifyGridCellContent("", 1, "Company code", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS_OCC1)
Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 2, "Posting Key", 0,DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "Posting Key", 0,DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 5, "Posting Key", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("", 6, "Posting Key", 0,DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("", 7, "Posting Key", 0,DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BSCHL)


Call VerifyGridCellContent("", 1, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("", 3, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 5, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("", 6, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)
Call VerifyGridCellContent("", 7, "Account", 0, DT_MB90_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_KTONR)

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


