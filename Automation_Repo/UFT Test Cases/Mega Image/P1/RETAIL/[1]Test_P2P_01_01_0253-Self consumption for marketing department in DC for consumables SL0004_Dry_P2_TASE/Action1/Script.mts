
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

gstrTestCaseName = "Test_P2P_01_01_0253-P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
'
'''''--------TransactionCode-MB51----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextBox("Order","VBAK-VBELN",0,DT_VA03_0102_ORDER,False)

Call ClickButton("Display document flow   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SelectNodeGuiTree(0,"#1;#1")
Call DoubleClick()
Call ClickButton("Display document   \(F8\)",False)
Call GetTextBoxValue("LIKP-VBELN",0,"DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT",False)
Call TakeScreenShot

Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call VerifyStatusBar("Self Consumption "&DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT&" has been saved")

Call SelectNodeGuiTree(0,"#1;#2")
Call DoubleClick()
Call ClickButton("Display document   \(F8\)",False)
Call GetTextBoxValue("LIKP-VBELN",0,"DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OCC1_OUTPUT",False)

Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False)
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call VerifyStatusBar("Self Consumption "&DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OCC1_OUTPUT&" has been saved")

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Display document flow   \(Shift\+F5\)",False)
Call SelectNodeGuiTree(0,"#1;#1;#1")
Call DoubleClick()
Call GetGridContent("Self.*", 0, "Doc.no.", 1,"Number of Preceding Document", DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT, "DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
Call SelectNodeGuiTree(0,"#1;#2;#1")
Call DoubleClick()
Call GetGridContent("Self.*", 0, "Doc.no.", 1, "Number of Preceding Document", DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OCC1_OUTPUT, "DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OCC1_OUTPUT")

Call SetTcode(DT_VA03_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT,False)
Call PressEnter()     
Call PressEnter() 
Call TakeScreenShot
Call FocusTextBoxByIndex("Material Description","MSEG-MATNR",0,False)
Call ClickButton("Choose   \(F2\)",False)
Call ClickButton("Output   \(Shift\+F2\)",False)



Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_VA03_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_VA03_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
'''Call ClickButton("Back   \(F3\)",False)
'''Call ClickButton("Back   \(F3\)",False)
'While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Article Doc\.","name:=RM07M-MBLNR","Index:=0").Exist(5)=False 
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
'Wend
Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OCC1_OUTPUT,False)
Call PressEnter()     
Call PressEnter() 
Call TakeScreenShot
Call FocusTextBoxByIndex("Material Description","MSEG-MATNR",0,False)
Call ClickButton("Choose   \(F2\)",False)
Call ClickButton("Output   \(Shift\+F2\)",False)
wait 5

Call VerifyTableCellContent(1,"Status", "SAPDV70ATC_NAST3", DT_VA03_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0_OCC1)
Call VerifyTableCellContent(1,"Output Type", "SAPDV70ATC_NAST3", DT_VA03_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0_OCC1)

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


