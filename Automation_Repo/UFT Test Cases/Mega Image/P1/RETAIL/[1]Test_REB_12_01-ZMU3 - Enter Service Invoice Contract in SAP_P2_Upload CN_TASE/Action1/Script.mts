
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_12_01-ZMU3 - Enter Service Invoice Contract in SAP_P2_Upload CN
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_12_01-ZMU3 - Enter Service Invoice Contract in SAP_P2_Upload CN
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




gstrTestCaseName = "Test_REB_12_01-ZMU3 - Enter Service Invoice Contract in SAP_P2_Upload CN"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_REB_12_01-ZMU3 - Enter Service Invoice Contract  in SAP_P2_Upload CN.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
''SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()
'
Call SelectCheckBox("P_TEST","0","OFF",False)

Call SetTextBox("Purchasing Organization","P_EKORG",0,DT_ZMDRB_MASS_UPL_DISC_1000_PURCHASING_ORGANIZATION,False)
Call SetTextBox("Company Code","P_BUKRS",0,DT_ZMDRB_MASS_UPL_DISC_1000_COMPANY_CODE,False)
Call SetTextBox("Date","P_WFDAT",0,DT_ZMDRB_MASS_UPL_DISC_1000_DATE,False)
Call SetTextBox("Filename","P_FILE",0,DT_ZMDRB_MASS_UPL_DISC_1000_FILENAME,False)
Call ClickButton("Execute   \(F8\)",False)
wait 5
Call TakeScreenshot()
Call VerifyGridCellContent("", 1, "Status", 0, DT_ZMDRB_MASS_UPL_DISC_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)
Call ClickCellGuiGrid("",0,"Payment document",1,"","",False)
'''Call ActivateCellGuiGridByRefVal("",0,"Document Date",DT_ZMDRB_MASS_UPL_DISC_1000_DATE,"Payment document",False)

Call SelectMenuBar("Payment Document;Change")
Call ClickButton("Release to Accounting   \(Ctrl\+F7\)",False)
wait 5


Call ClickButton("Accounting   \(F6\)",False)
'''Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)


Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call VerifyGridCellContent("", 1, "BSCHL", 0,DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)


Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





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




