



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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_ChangeCustomerCentrally_XD02_DeleteVAT 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_POST_ChangeCustomerCentrally_XD02_DeleteVAT"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_PFA_ChangeCustomerCentrally_XD02_171019_164356_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------XD02 -----------------------------------
Call SendKey("{F4}")

Call SelectTab("G_SELONETABSTRIP","T: Customers by Tax Information",True)

If VerifyTextBoxContent("Tax Number 2","G_SELFLD_TAB-LOW","","",True) =False Then
  Call ClickButtonIfExist("Cancel   \(F12\)",True)	
End If 

Call SetTextbox("Tax Number 2","G_SELFLD_TAB-LOW","",DT_XD02_0220_TAX_NUMBER_2,True) 
Call SetTextbox("VAT Registration No\.","G_SELFLD_TAB-LOW","",DT_XD02_0220_VAT_REGISTRATION_NO,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Copy   \(Enter\)",True)

Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD02_7101_COMPANY_CODE,True) 
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD02_7101_SALES_ORGANIZATION,True)
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD02_7101_DISTRIBUTION_CHANNEL,True) 
Call SetTextbox("Division","RF02D-SPART","",DT_XD02_7101_DIVISION,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SelectTab("TABSTRIP100","Control Data",False)
Call SetTextbox("Tax Number 2","KNA1-STCD2","","",False) 
Call SetTextbox("VAT Reg\. No\.","KNA1-STCEG","","",False) 
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()


'//-----------------------------------XD02 -----------------------------------
Call SendKey("{F4}")
If VerifyTextBoxContent("Tax Number 2","G_SELFLD_TAB-LOW","","",True) =False Then
  Call ClickButtonIfExist("Cancel   \(F12\)",True)	
End If 
Call SelectTab("G_SELONETABSTRIP","T: Customers by Tax Information",True)

Call SetTextbox("VAT Registration No\.","G_SELFLD_TAB-LOW","",DT_XD02_0220_VAT_REGISTRATION_NO_OCC1,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Copy   \(Enter\)",True)

Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD02_7101_COMPANY_CODE_OCC1,True) 
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD02_7101_SALES_ORGANIZATION_OCC1,True)
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD02_7101_DISTRIBUTION_CHANNEL_OCC1,True) 
Call SetTextbox("Division","RF02D-SPART","",DT_XD02_7101_DIVISION_OCC1,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SelectTab("TABSTRIP100","Control Data",False)
Call SetTextbox("VAT Reg\. No\.","KNA1-STCEG","","",False) 
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
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







