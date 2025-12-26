

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_09.06.05.17.011 Other minor expenses_Movement Type FE11_Retail_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.17.011 Other minor expenses_Movement Type FE11_Retail_R1E_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------/POSDW/MON0 -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
''''''''''''''''

call SelectRadioButton("X_AISEL","All items",false)

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",convertdate(DT_FAGLL03_1000_POSTING_DATE),False)
call ClickButton("Custom Selections   \(Ctrl\+F1\)",fALSE)
'var = SAPGuiSession("Session_2").SAPGuiWindow("Define Filter Criteria").SAPGuiGrid("Column Set").FindRowByCellContent("Column Name","Text")
'
'nRowno = SAPGuiSession("Class Name:=SAPGuiSession").SAPGuiWindow("text:=Define Filter Criteria").SAPGuiGrid("title:=Column Set").FindRowByCellContent("Column Name","Text")

' ActivateNodeGuiTree(treeIndex, itemPath)
call ActivateNodeGuiTree(1,"General Ledger Line Items;Profit Center")
Call SetTextbox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
call ClickButton("Save   \(Ctrl\+S\)",false)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()


''call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",fALSE)
''' SelectCellGuiGrid(gridTitle, gridIndex, Rowno, columnName, blnIsItPopup)
'''msgbox SAPGuiSession("Class Name:=SAPGuiSession","activewindow:=Define Filter Criteria").Exist
'''nRowno = SAPGuiSession("Class Name:=SAPGuiSession","activewindow:=Define Filter Criteria").SAPGuiWindow("text:=Define Filter Criteria").SAPGuiGrid("title:=Column Set").FindRowByCellContent("Column Name","Text")
''nRowno = SAPGuiSession("Session_2").SAPGuiWindow("Define Filter Criteria").SAPGuiGrid("Column Set").FindRowByCellContent("Column Name","Text")
''call SelectCellGuiGrid("Column Set",1,nRowno,"Column Name",True)
''call ClickButton("Add Filter Criterion \(F7\)",fALSE)
''call ClickButton("Define Filter Values",fALSE)
''Call SetTextbox("Text","%%DYN001-LOW","",DT_FAGLL03_1105_TEXT,False)
''call ClickButton("Execute   \(Enter\)",fALSE)
''Call TakeScreenSHot()


Call  SelectColumnGuiGrid("", 1, "Text", False)
Call TakeScreenSHot()
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenSHot()
Call SetTextboxNoName("Text","",DT_FAGLL03_1105_TEXT,False)
call ClickButton("Execute   \(Enter\)",False)
Call TakeScreenshot()

call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",fALSE)
call ClickButton("Call Up Document Overview   \(F9\)",fALSE)
' SAPGuiSession("Session_2").SAPGuiWindow("Display Document: Data").SAPGuiGrid("GridViewCtrl").
' msgbox SAPGuiSession("Session_2").SAPGuiWindow("Display Document: Data").SAPGuiGrid("GridViewCtrl").GetCellData(1,2)
'PostingKey = SAPGuiSession("Session_2").SAPGuiWindow("Display Document: Data").SAPGuiGrid("GridViewCtrl").GetCellData(1,2)
'ProfitCenter = SAPGuiSession("Session_2").SAPGuiWindow("Display Document: Data").SAPGuiGrid("GridViewCtrl").GetCellData(1,15)
call VerifyGridCellContent("",1,"Posting Key",2,PostingKey)
'call VerifyGridCellContent("",1,"Posting Key",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Profit Center",2,ProfitCenter)
'call VerifyGridCellContent("",1,"Posting Key",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)


'Call SetTextbox("Reference","BKPF-XBLNR","",DT_FAGLL03_0750_CHECK_TEXT_OF_REFERENCE,False)

'Transfer   \(Enter\)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)

Call LogOff()
Call FinalStatus ()











''''''''''''''''''''''''
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


