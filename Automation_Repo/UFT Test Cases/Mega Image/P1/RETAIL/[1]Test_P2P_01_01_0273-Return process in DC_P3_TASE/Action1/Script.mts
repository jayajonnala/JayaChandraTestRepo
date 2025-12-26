
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P3
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




gstrTestCaseName = "Test_P2P_01_01_0273_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0273-Return process in DC_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

''''''''MB51

Call SetTcode(DT_SAPTRANSACTIONCODE)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Reference","XBLNR-LOW",0,DT_MB51_1000_REFERENCE,False)

Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)

Call  SelectRowGuiGrid("", 0, "Reference", DT_MB51_1000_REFERENCE, False)

Call DoubleClick()

Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Output",False)

Call ClickButton("OK_NAST_SHOW",False)

Call TakeScreenShot()

'Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_MB51_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
'Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_MB51_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_MB51_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Sort order","PM_NSORT",0,DT_MB51_1000_SORT_ORDER,False)

Call SetTextBox("Processing Mode","PM_VERMO",0,DT_MB51_1000_PROCESSING_MODE,False)

Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_YEAR,False)

Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MB51_1000_ARTICLE_DOCUMENT,False)

Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel(0, "ON", False)

Call TakeScreenShot()

Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call TakeScreenShot()
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_MB51_0100_OKCD_OCC1)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Purch\. Organization","EKORG-LOW",0,DT_MB51_1000_PURCH_ORGANIZATION,False)

Call SetTextBox("Purchasing Document","EBELN-LOW",0,DT_MB51_1000_PURCHASING_DOCUMENT,False)

Call SetTextBox("Movement type","S_BWART-LOW",0,DT_MB51_1000_MOVEMENT_TYPE,False)

Call SetTextBox("Company Code","S_BUKRS-LOW",0,DT_COMPANY_CODE,False)

Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel(0, "ON", False)

Call SelectCheckboxNoLabel(1, "ON", False)

Call TakeScreenShot()

Call ClickButton("Send notification   \(Shift\+F2\)",False)

Call TakeScreenShot()

Call ClickButton("Cancel   \(F12\)",False)

Call LogOff()

Call FinalStatus()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





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




