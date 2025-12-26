
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P2     

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


gstrTestCaseName = "Test_P2P_01_01_0253_Fresh_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0253-Self consumption for marketing department in DC for consumables SL0004_Fresh_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextBox("Shipping point","LIKP-VSTEL",0,DT_SHIPPING_PONT,False)
Call SetTextBox("Selection date","LV50C-DATBI",0,ConvertDate(DT_SELECTION_DATE),False)
Call SetTextBox("Order","LV50C-VBELN",0,DT_VA03_0102_ORDER,False)
Call PressEnter()
Call ClickButton("btn\[11\]",False)
Call TakeScreenShot()
Call GetStatusBar("item2","DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT")
Call SetTcode(DT_VL03N)
Call PressEnter()
Call SetTextBox("Outbound Delivery","LIKP-VBELN",0,DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT,False)
Call PressEnter()
Call ClickButton("btn\[25\]",False)
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call GetStatusBar("item3","DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
'Call VerifyStatusBar("Self Consumption "&DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT&" has been saved")
Call VerifyStatusBar("Self Consumption "&DT_VA03_1502_CHECK_TEXT_OF_OUTBOUND_DELIV_OUTPUT&" saved, article document "&DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT&" created")
Call TakeScreenshot()
Call SetTcode(DT_VA03)
Call PressEnter()
Call SetTextBox("Order","VBAK-VBELN",0,DT_ORDER,False)
Call ClickButton("Display document flow   \(Shift\+F5\)",False)
Call TakeScreenShot()
Call SelectNodeGuiTree(0,"#1;#1;#1")
Call DoubleClick()

Call GetGridContent("Self.*", 0, "Doc.no.", 1, "Item (SD)", "1", "DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")

Call SetTcode(DT_VA03_0100_OKCD)
Call PressEnter()



Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("Article Doc\.","RM07M-MBLNR",0,DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM,False)
Call TakeScreenShot()
Call PressEnter()
wait 5
Call PressEnter()

Call ClickButton("Details from Item\.\.\.   \(Shift\+F8\)",False)
Call SetTextBox("Art\. Doc\.Item","RM07M-ZEILE",0,1,True)
Call PressEnter()
Call ClickButton("btn\[14\]",False)
Call VerifyTableCellContent(1,"Status", "SAPDV70ATC_NAST3", DT_VA03_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1,"Output Type", "SAPDV70ATC_NAST3", DT_VA03_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

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


