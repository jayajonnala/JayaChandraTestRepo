

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRI_01_038-ZCC1 calculated for PO_ZHC5_GR   
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

gstrTestCaseName = "Test_S2A_PRI_01_038-ZCC1 calculated for PO_ZHC5_GR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRI_01_038-ZCC1 calculated for PO_ZHC5_GR.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''''-----Login----------'''''''
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-ME22N ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME22N_0003_PUR_ORDER,True) 
Call PressEnter()
Call ClickButtonIfExist("No",False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL", "Texts", False)
Call SelectTab("HEADER_DETAIL", "Org. Data", False)
''If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
''Call ClickButtonIfExist("Display/Change   \(F7\)",False)
''End If
''Call TakeScreenShot
''Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("HEADER_DETAIL","Conditions", False)
Call TakeScreenShot
Call FindRowNumber("SAPLV69ATCTRL_KONDITIONEN", "CnTy", "ZNAV", "DT_ROW_NO_OUTPUT")
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN", "CnTy", DT_ROW_NO_OUTPUT+1, "<NA>", "<NA>", DT_ME22N_6201_TABLECELL_CNTY_4, False)
Call SetTableData("SAPLV69ATCTRL_KONDITIONEN", "Amount", DT_ROW_NO_OUTPUT+1, "<NA>", "<NA>", DT_ME22N_6201_TABLECELL_AMOUNT_4, False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot
Call GetStatusBar("Text","DT_ME22N_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_ME22N_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT)


''''''''--------TransactionCode-ME29n----------''''

Call SetTcode(DT_ME22N_0014_OKCD)     
Call PressEnter() 

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME22N_0003_PUR_ORDER,True) 
Call PressEnter()
Call ClickButtonIfExist("No",False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL", "Texts", False)
Call SelectTab("HEADER_DETAIL", "Org. Data", False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call TakeScreenShot
Call SelectTab("HEADER_DETAIL", "Release strategy", False)
Call TakeScreenShot
Call ClickCellGuiGrid("", 0, "Release options", 1, "<NA>", "<NA>", False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot
Call GetStatusBar("Text","DT_ME22N_0014_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call VerifyStatusBar(DT_ME22N_0014_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT)
'
'''''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_ME22N_0014_OKCD_OCC1)     
Call PressEnter() 

Call SetComboByKey("GODYNPRO-ACTION", DT_ME22N_0010_GODYNPROACTION)
Call TakeScreenShot
Call SetComboByKey("GODYNPRO-REFDOC", DT_ME22N_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_ME22N_0003_PUR_ORDER, False)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME22N_0110_DELIVERY_NOTE,False)
Call PressEnter()
Call ClickButtonIfExist("BUTTON_DETAIL",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("Text","DT_ME22N_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_ME22N_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT)


Call LogOff()

Call FinalStatus()






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




