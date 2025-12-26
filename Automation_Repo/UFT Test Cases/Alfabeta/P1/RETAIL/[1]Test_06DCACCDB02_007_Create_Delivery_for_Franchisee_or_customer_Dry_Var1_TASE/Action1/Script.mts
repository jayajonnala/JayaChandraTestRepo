

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCACCDB02_007_Create_Delivery_for_Franchisee_or_customer_Dry_Var1
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

gstrTestCaseName = "Test_06DCACCDB02_007_Create_Delivery_for_Franchisee"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_06DCACCDB02_007_Create_Delivery_for_Franchisee_or_customer_Dry_Var1_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  VL10G------------------------------------------------
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
Call SetTextbox("CalcRuleDefltDlvCrDt","P_LERUL","","",False)
Call TakeScreenShot()
Call ClickButton("%_ST_VBELN_%_APP_%-VALU_PUSH",False) 
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_VL10G_3010_TABLECELL_SINGLE_VALUE_0,True)
Call  ClickButton("Copy   \(F8\)",True) 
Call TakeScreenShot()

Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call ClickButton("%_ST_EBELN_%_APP_%-VALU_PUSH",False) 
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_VL10G_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call  ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()

Call ClickCellGuiGrid("",0,"Originating Document",1,"","",False)
Call VerifyTableCellContent(1,"Storage Location","SAPMV45ATCTRL_U_ERF_AUFTRAG",DT_STORAGE_LOC) 
Call ClickButton("Back   \(F3\)",False) 

Call SelectRowGuiGrid("",0,"Originating Document",DT_VL10G_3010_TABLECELL_SINGLE_VALUE_0,False)
Call ClickButton("Create Delivery in Background   \(Shift\+F7\)",False) 
Call TakeScreenShot()

'Call SelectColumnGuiGrid("",0,"Sales Document",False)
Call SelectColumnGuiGrid("",0,"SD Document",False)
Call TakeScreenShot()
Call ClickButton("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)
Call TakeScreenShot()
'Call GetGridContent("",0,"Sales Document",1,"<NA>","<NA>","DT_OUTB_DEL_OUTPUT")
Call GetGridContent("",0,"SD Document",1,"<NA>","<NA>","DT_OUTB_DEL_OUTPUT")
Call TakeScreenShot()

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




