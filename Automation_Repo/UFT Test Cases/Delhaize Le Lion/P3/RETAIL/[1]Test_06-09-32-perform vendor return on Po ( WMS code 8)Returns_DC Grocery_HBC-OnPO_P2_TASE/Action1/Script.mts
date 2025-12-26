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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06-09-32-perform vendor return on Po ( WMS code 8)Returns_DC Grocery_HBC-OnPO_P2_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06-09-32-retu_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-05-Send PO to WMS_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call Login(DT_SAPUSER,DT_SAPPASSWORD)
'Call PressEnter() 

''--------TransactionCode-ME23N----------''''

'' ME23N
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call SelectTab("ITEM_DETAIL","Purchase Order History",False)

'Call ClickButton("Next item",False)
 
Call SelectColumnGuiGrid("","0","Movement type",False)
Call TakeScreenShot

wait 3
' ClickContextButtonToolBar(buttonName, toolbarIndex)
'Call ClickContextButtonToolBar("&MB_FILTER",0)
Call Click204ButtonToolBar("&MB_FILTER", 0)


'Call ClickButtonToolBar("&MB_FILTER",1)

Call SetTextboxPopupIfExist("%%DYN001-LOW","Movement type",DT_ME23N_1105_MOVEMENT_TYPE)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("",1,"Movement type","0",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
Call VerifyGridCellContent("",1,"Quantity","0",DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MENGE)
Call TakeScreenShot
Call ClickCellGuiGrid("","0","Article Document",1,"","",False)
''Call SelectCellGuiGrid("","0",1,"Article Document",False)
Call VerifyTextBoxContent("HeaderText","GOHEAD-BKTXT","0",DT_ME23N_0110_CHECK_TEXT_OF_HEADERTEXT,False)
' ClickCellGuiGrid(gridTitle, gridIndex, columnName, rowNumber, columnNameRef, gridValRef, blnIsItPopup)
Call TakeScreenShot


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



