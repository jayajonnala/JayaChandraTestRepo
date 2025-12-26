

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


gstrTestCaseName = "Test_09.06.05.12.02 Type F907_Retail_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.12.02 Receipt Payment with Credit Card _Movement Type F907_Retail_R1E_TASE.xls"


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

'//-----------------------------------/FBL5N -----------------------------------

call SelectRadioButton("X_AISEL","All items",false)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)   
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)   
Call SetTextbox("to","SO_BUDAT-HIGH","",convertdate(DT_FBL5N_1000_POSTING_DATE),False)
wait(5)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
call ClickButtonIfExist("Continue   \(Enter\)",true)
call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",fALSE)
call ClickButtonToolBar("&FIND",0)
call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBL5N_0841_SEARCH_TERM,true)
Call TakeScreenShot()
call ClickButton("OK   \(Enter\)",true)
call ClickButton("Cancel   \(F12\)",true)
Call PressEnter()
call ClickButtonIfExist("Transfer   \(Enter\)",true)
'''''

Call SelectColumnGuiGrid("", 0, "Text", False)
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",fALSE)
Call TakeScreenSHot()
Call SetTextbox("Text","%%DYN001-LOW","",DT_FBL5N_1105_TEXT,False)
call ClickButton("Execute   \(Enter\)",fALSE)
call ClickButton("Display Document   \(Shift\+F2\)",fALSE)

call ClickButton("Call Up Document Overview   \(F9\)",fALSE)

Call VerifyGridCellContent("", 1, "BSCHL", "", DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", "", DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

call VerifyGridCellContent("",1,"PRCTR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
call VerifyGridCellContent("",2,"PRCTR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

call VerifyGridCellContent("",1,"KTONR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"KTONR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

'
'Call GetGridContent("", 0, "Amount", 1, "", "", "DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OUTPUT")
'Call GetGridContent("", 0, "Amount", 2, "", "", "DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OUTPUT")
'
call VerifyGridCellContent("",1,"AZBET","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OUTPUT)
call VerifyGridCellContent("",2,"AZBET","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OUTPUT)

call VerifyGridCellContent("",1,"SGTXT","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
call VerifyGridCellContent("",2,"SGTXT","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call LogOff()
Call FinalStatus ()
''''''''''''''''''''''''''''''''''''''''
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


