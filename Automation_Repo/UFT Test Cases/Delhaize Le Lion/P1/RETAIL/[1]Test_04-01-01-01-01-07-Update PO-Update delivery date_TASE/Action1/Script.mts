
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04-01-01-01-01-07-Update PO-Update delivery date
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

gstrTestCaseName = "Test_04-01-01-01-01-07-Update PO-Update delivery date"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-07-Update PO-Update delivery date_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'--------------------------------------------  ME21N------------------------------------------------

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME21N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL", "Partners", False)
Call TakeScreenShot


Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-SUPERFIELD",0,DT_ME21N_1105_CHECK_TEXT_OF_VENDOR,False)
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call VerifyTableCellContent(1,"Article", "SAPLMEGUITC_1211", DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(1,"PO Quantity", "SAPLMEGUITC_1211", DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0)
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Deliv. Date", "SAPLMEGUITC_1211", ConvertDate(DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_0))
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False)    
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Save",True) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'''''Added 80 and 81 by KGARA on 28/7/22'''''
Call GetStatusBar("item2","DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call  VerifyStatusBar("Standard PO Retail "&DT_ME21N_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" changed")
wait(5)
Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Deliv. Date", "SAPLMEGUITC_1211", ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0))

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


