'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Upload AR Accounting Document_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th Feb
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Upload AR Accounting Document_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Upload AR Accounting Document_p2.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call TakeScreenShot()
Call PressEnter()  
'
'----------------------Tcode SM35----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
'___________________________________________________________SM35_______________________________________________________________

Call SetTextboxNoLabel("D0100-MAPN", 0, DT_SESSION, False)
Call SetTextbox("From:", "D0100-VON", 0, DT_FROM, False)
Call SetTextbox("Created by:", "D0100-CREATOR", 0, DT_CREATED_BY, False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call GetTableCellData("SAPMSBDC_CCTC_APQI", "Status", 1, "Session Name", "GTE089", "DT_STATUS_OUTPUT", False)
If DT_STATUS_OUTPUT<>"STA_OK" Then
	'select row
Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name","GTE089",False)
Call ClickButton("Process session   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call SelectRadioButtonIfPopupExists("D0300-ERROR","Display errors only")
Call ClickButton("Process   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()
Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
End If


'Capture the screenshot
call TakeScreenShot()
Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name","GTE089",False)
'Capture the screenshot
call TakeScreenShot()
Call ClickButton("Analyze session   \(F2\)",False)
'Capture the screenshot
call TakeScreenShot()
Call SelectTab("TAB_DYNPRO"," Dynpros",False)
'Capture the screenshot
call TakeScreenShot()
Call SelectTab("TAB_DYNPRO"," Transactions",False)
'Capture the screenshot
call TakeScreenShot()
'processed checkpoints
call VerifyTableCellContent(DT_INDEX,"Status","RSBDC_ANALYSETC_TCODES",Lcase(DT_SM35_0200_CHECK_TEXT_OF_TABLECELL_STATUS_0))
call VerifyTableCellContent(DT_INDEX1,"Status","RSBDC_ANALYSETC_TCODES",Lcase(DT_SM35_0200_CHECK_TEXT_OF_TABLECELL_STATUS_1))
'---------------------------------------Document 1 / Index 1 -----------------------------------------------
Call SelectRowGuiTableByRow("RSBDC_ANALYSETC_TCODES",DT_INDEX,False)
'Call SelectCellGuiTable("RSBDC_ANALYSETC_TCODES","Index","Index",DT_INDEX,False)
Call SendKey("{F2}")
'Capture the screenshot
Call TakeScreenShot()
'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&Replace((DT_DATE),"/","."),False)
Wait(1)
Call TakeScreenShot()
Call SelectRadioButton("RB-PRO_TCODE","Transaction",False)
Call TakeScreenShot()
'fetch oupput mesage row nummber
Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL","T","S","DT_GET_DOC_ROW_output")
'loading datasheet again
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'fetch oupput mesage
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",DT_GET_DOC_ROW,"T","S","DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_0__output",False)
'---------------------------------------Document 2 / Index 2 -----------------------------------------------
Call SelectTab("TAB_DYNPRO"," Transactions",False)
'Capture the screenshot
call TakeScreenShot()
'Call SelectRowGuiTableByRow("RSBDC_ANALYSETC_TCODES",DT_INDEX1,False)
Call SelectCellGuiTable("RSBDC_ANALYSETC_TCODES","Index","Index",DT_INDEX1,False)
Call SendKey("{F2}")
'Capture the screenshot
Call TakeScreenShot()
'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&Replace((DT_DATE),"/","."),False)
Wait(1)
Call TakeScreenShot()
Call SelectRadioButton("RB-PRO_TCODE","Transaction",False)
Call TakeScreenShot()
'fetch oupput mesage row nummber
Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL","T","S","DT_GET_DOC_ROW_1_output")
'loading datasheet again
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'fetch oupput mesage
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",DT_GET_DOC_ROW_1,"T","S","DT_SM35_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2_output",False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()




