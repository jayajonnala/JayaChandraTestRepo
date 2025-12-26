
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_34_01-ZME3 - Enter Extras SI contract in SAP_P2_Upload SI_TASE      
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


gstrTestCaseName = "Test_REB_34_01-ZME3_P2_Upload SI_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_REB_34_01-ZME3 - Enter Extras SI contract in SAP_P2_Upload SI.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''--------TransactionCode-WE02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot


Call SelectCheckBox("P_TEST","0","OFF",False)

Call SetTextBox("Purchasing Organization","P_EKORG",0,DT_ZMDRB_MASS_UPL_DISC_1000_PURCHASING_ORGANIZATION,False)
Call SetTextBox("Company Code","P_BUKRS",0,DT_ZMDRB_MASS_UPL_DISC_1000_COMPANY_CODE,False)
Call SetTextBox("Date","P_WFDAT",0,ConvertDate(DT_ZMDRB_MASS_UPL_DISC_1000_DATE),False)
Call SetTextBox("Filename","P_FILE",0,DT_ZMDRB_MASS_UPL_DISC_1000_FILENAME,False)
Call ClickButton("Execute   \(F8\)",false)

Call VerifyGridCellContent("", 1, "Status", 0, DT_ZMDRB_MASS_GRIDCELL_0_STATUS)
''''Call ActivateCellGuiGridByRefVal("",0,"Document Date",DT_ZMDRB_MASS_UPL_DISC_1000_DATE,"Payment document",False)
''Call ActivateCellGuiGridByRefVal("",0,"Payment type",DT_PAYMENT_TYPE,"Payment document",False)
''Call ClickCellGuiGrid("", 0, "Payment document", 1, "Payment type", DT_PAYMENT_TYPE, False)
Call GetGridContent("", 0, "Payment document", 1, "Payment type", DT_PAYMENT_TYPE, "DT_PAYMENT_DOC_OUTPUT")
Call TakeScreenShot
Call SetTcode(DT_SAPTCODE_1)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextBox("Payment document","RWLF1-WBELN_ZR",0,DT_PAYMENT_DOC_OUTPUT,False)
Call ClickButton("Release to Accounting   \(Ctrl\+F7\)",False)
Call TakeScreenShot
Call SetTcode(DT_SAPTCODE_2)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextBox("Payment document","RWLF1-WBELN_ZR",0,DT_PAYMENT_DOC_OUTPUT,False)
Call ClickButton("Accounting   \(F6\)",False)
'Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Doc. Number", True)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)
Call VerifyGridCellContent("", 1, "BSCHL", 0,DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_ZMDRB_MASS_UPL_DISC_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
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



