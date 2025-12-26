

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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_06FIACC_017_Consumption"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_06FIACC_017_Self_Consumption_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

'//-----------------------------------VF03 -----------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Billing document","VBRK-VBELN","",DT_VF03_0101_BILLING_DOCUMENT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Accounting overview   \(Shift\+F4\)",False)
Call TakeScreenShot

Call DoubleClickGuiGridCell("Documents in Accounting",0,1,"Doc. Number",True)
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call ClickButton("Display Document Header   \(F5\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Document Type","BKPF-BLART",0,DT_VF03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
Call TakeScreenShot
Call ClickButton("Continue/Confirm   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("",1,"Alternative Account No.",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("",2,"Alternative Account No.",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("",3,"Alternative Account No.",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LOKKT)
Call TakeScreenShot

Call ClickContextButtonToolBar("&MB_VARIANT",0)
Call TakeScreenShot
Call SelectMenuIdToolBar("&COL0",0)
Call TakeScreenShot

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
Call TakeScreenShot
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call TakeScreenShot

Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call GetTextboxValue("BKPF-BELNR",0,"DT_VF03_0200_GRIDCELL_0_DOC_NUMBER_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_VF03_0200_GRIDCELL_0_DOC_NUMBER_OUTPUT",DT_VF03_0200_GRIDCELL_0_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

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


