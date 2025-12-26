
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_01_01-ZMU1 - Enter CN Contract in SAP_P3_settlement
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
'.................Test Script Name :Test_REB_01_01-ZMU1 - Enter CN Contract in SAP_P3_settlement
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_REB_SI_Fixed Amount_Auto_Invoice_Posting"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_REB_01_01-ZMU1 - Enter CN Contract  in SAP_P3_settlement.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call EndDateof445PeriodByDate(DT_DATE,"DT_ENDING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()


Call SetTextBox("Condition Contract","S_NUM-LOW",0,DT_WB2R_SV_1000_CONDITION_CONTRACT,False)
Call SetTextBox("Condition Contract Type","S_CTRTP-LOW",0,DT_CONTRACT_TYPE,False)
Call SetTextBox("Company Code","S_BUKRS-LOW",0,DT_COMPANY_CODE,False)
Call SetTextBox("Purch. Organization","S_EKORG-LOW",0,DT_PURCHASING_ORG,False)
Call SetTextBox("Valid from","P_DATEFR",0,ConvertDate(DT_VALID_FROM),False)
Call SetTextBox("Valid to","P_DATETO",0,ConvertDate(DT_VALID_TO),False)
Call SetTextBox("Settlement Date From","P_SDATFR",0,ConvertDate(DT_WB2R_SV_1000_SETTLEMENT_DATE),False)
Call SetTextBox("Settlement Date To","P_SDATTO",0,ConvertDate(DT_WB2R_SV_1000_TO),False)
Call SetTextBox("Posting Date","P_BUDAT",0,ConvertDate(DT_WB2R_SV_1000_POSTING_DATE),False)
Call SetTextBox("Document Date","P_BLDAT",0,ConvertDate(DT_WB2R_SV_1000_DOCUMENT_DATE),False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call GetGridContent("",0,"Document number",1,"Contract",DT_WB2R_SV_1000_CONDITION_CONTRACT,"DT_DOCUMENT_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1)     
Call PressEnter()
Call SetTextBox("Payment document","RWLF1-WBELN_ZR",0,DT_WB2R_SV_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
Call TakeScreenShot()
Call ClickButton("Accounting   \(F6\)",False)
Call TakeScreenShot()
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




