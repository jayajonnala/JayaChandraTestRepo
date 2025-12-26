

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_AB_023_ Create ZCXT Retail Customer_Other Debtor 
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If


gstrTestCaseName = "Test_MD_AB_023_ Create ZCXT Retail Customer_Other Debtor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_AB_023_ Create ZCXT Retail Customer_Other Debtor_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------XD01 -----------------------------------

Call SetTextbox("Customer","RF02D-KUNNR","",DT_XD01_7100_CUSTOMER,True)  
Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD01_7100_COMPANY_CODE,True)   
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD01_7100_SALES_ORGANIZATION,True) 
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD01_7100_DISTRIBUTION_CHANNEL,True)     
Call SetTextbox("Division","RF02D-SPART","",DT_XD01_7100_DIVISION,True) 
Call TakeScreenShot()

Call SetTextboxNoLabel("RF02D-REF_KUNNR",0,DT_XD01_7100_RF02DREF_KUNNR,True)
Call SetTextbox("Company code","RF02D-REF_BUKRS","",DT_XD01_7100_COMPANY_CODE_OCC1,True)  
Call SetTextbox("Sales organization","RF02D-REF_VKORG","",DT_XD01_7100_SALES_ORGANIZATION_OCC1,True)  
Call SetTextbox("Distribution channel","RF02D-REF_VTWEG","",DT_XD01_7100_DISTRIBUTION_CHANNEL_OCC1,True)  
Call SetTextbox("Reference division","RF02D-REF_SPART","",DT_XD01_7100_REFERENCE_DIVISION,True)  
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call SelectTab("TABSTRIP100","Partner Functions",False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_CUSTOMER_OUTPUT")
Call VerifyStatusBar("Customer "& DT_CUSTOMER_OUTPUT &" was created for sales area GS02 20 01")


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







