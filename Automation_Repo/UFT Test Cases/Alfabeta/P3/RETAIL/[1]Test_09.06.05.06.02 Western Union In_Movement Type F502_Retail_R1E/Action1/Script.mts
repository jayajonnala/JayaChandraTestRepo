

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

gstrTestCaseName = "Test_09.06.05.06.02 Western Union In_Movement Type F502_Retail_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.06.01 Western Union Out_Movement Type F502_R1E_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------FBL5N -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
call SelectRadioButton("X_AISEL","All items",false)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)   
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)   
Call SetTextbox("Posting date","SO_BUDAT-LOW","",ConvertDAte(DT_FBL5N_1000_POSTING_DATE),False)
wait(5)
call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenShot()
call ClickButton("Hide dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
wait(2)
call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenShot()

call ClickButtonIfExist("Continue   \(Enter\)",True)
'wait(20)
call ClickButtonToolBar("&FIND",0)
call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBL5N_0841_SEARCH_TERM,true)
'' SetCombo(attachedTextOrComboName, comboValue)
'call SetCombo("Search Direction",DT_FBL5N_0620_GRIDCELL_39_COLUMN_NAME)
TakeScreenShot()
call ClickButton("OK   \(Enter\)",True)
call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot()
call ClickButton("Show Selected Fields \(F7\)",True)
'call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenShot()
call ClickButtonToolBar("&FIND",0)
call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBL5N_0841_SEARCH_TERM_OCC1,true)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0,DT_FBL5N_0841_FIND_ONLY_ENTIRE_WORD_OR_VALUE_OCC1,True)
TakeScreenShot()
call ClickButton("OK   \(Enter\)",True)
call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot()
call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
call ClickButtonToolBar("&FIND",0)
call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBL5N_0841_SEARCH_TERM,True)
call ClickButton("OK   \(Enter\)",True)
call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot()
call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Text","%%DYN001-LOW","",DT_FBL5N_1105_TEXT,True)
Call ClickButton("Execute   \(Enter\)",True)

''''Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
''''call ClickButtonToolBar("&FIND",0)
''''call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBL5N_0841_SEARCH_TERM_OCC1,True)
''''Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0,DT_FBL5N_0841_FIND_ONLY_ENTIRE_WORD_OR_VALUE_OCC1,True)
''''call ClickButton("OK   \(Enter\)",True)
''''call ClickButton("Cancel   \(F12\)",True)
''''Call TakeScreenShot()
''''call ClickButton("Add Filter Criterion \(F7\)",True)
''''Call ClickButton("Define Filter Values",True)
''''Call SetTextbox("Reference","%%DYN002-LOW","",DT_FBL5N_1105_TEXT_OCC1,True)
''''Call ClickButton("Execute   \(Enter\)",True)

call ClickButton("Display Document   \(Shift\+F2\)",False)
Call TakeScreenShot()
call ClickButton("Call Up Document Overview   \(F9\)",False)
'' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Text",0,DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE)

Call GetTextboxValue("BKPF-XBLNR",0,"DT_REF_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_REF_OUTPUT",DT_REF_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("",1,"Posting Key",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",1,"Account",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"Amount",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"Amount",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",1,"Profit Center",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",2,"Profit Center",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("",1,"Text",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("",2,"Text",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)


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


