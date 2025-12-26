

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

gstrTestCaseName = "Test_07.1.4.3.01 AB_P2_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.1.4.3.01 AB Store Self Consumption _Fuels_P2_R1E.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------/POSDW/MON0 -----------------------------------

Call SetTextbox("Site","WERKS-LOW","",DT_MB51_1000_SITE,False) 
Call SetTextbox("Posting Date","BUDAT-LOW","",ConvertDate(DT_MB51_1000_POSTING_DATE),False) 
Call SetTextbox("to","BUDAT-HIGH","",ConvertDate(DT_MB51_1000_TO),False) 
'Call SetTextbox("Reference","XBLNR-LOW","",DT_MB51_1000_REFERENCE,False) 
Call SetTextbox("Reference","XBLNR-LOW","","",False)
Call SelectRadioButton("RFLAT_L","Flat List",False)
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
Call SetTcode(DT_MB51_0500_OKCD) 
Call PressEnter() 
'Call SetTextbox("to","SD_SAKNR-HIGH","",DT_MB51_1000_TO_OCC1,False)   
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_MB51_1000_COMPANY_CODE,False)   
Call SetTextbox("to","SD_SAKNR-HIGH","",DT_MB51_1000_TO_OCC1,False)   
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_MB51_1000_GL_ACCOUNT,False)   
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_MB51_1000_POSTING_DATE_OCC1),False)   
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT),False) 
Call SelectRadioButton("X_AISEL","All Items",False)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",fALSE)
Call TakeScreenSHot()
' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
call ClickButtonIfExist("Call Up Document Overview   \(F9\)",false)
Call TakeScreenSHot()
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


