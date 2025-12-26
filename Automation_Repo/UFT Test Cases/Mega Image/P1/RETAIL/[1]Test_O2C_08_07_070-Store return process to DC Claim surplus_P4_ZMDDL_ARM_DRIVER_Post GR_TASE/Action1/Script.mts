
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_070-Store return process to DC Claim surplus_P4_ZMDDL_ARM_DRIVER_Post GR     
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


gstrTestCaseName = "Test_O2C_08_07_070_P4_ZMDDL_ARM_DRIVER_Post GR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_070-Store return process to DC Claim surplus_P4_ZMDDL_ARM_DRIVER_Post GR.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''--------TransactionCode-ZMDDL_ARM_DRIVER----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectRadioButton("P_SITE", "Store", False)
Call SetTextbox("Document Date","S_BLDAT-LOW","",ConvertDate(DT_ZMDDL_ARM_DRIVER_1000_DOCUMENT_DATE),False)
Call SetTextbox("to","S_BLDAT-HIGH","",ConvertDate(DT_ZMDDL_ARM_DRIVER_1000_TO),False)
Call SetTextbox("Delivery Note ID","S_LIFEX-LOW","",DT_ZMDDL_ARM_DRIVER_1000_DELIVERY_NOTE_ID,False)
Call TakeScreenShot
Call SelectRadioButton("P_GR", "Perform Goods Receipt", False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call SetTextbox("Delivery","SO_DLVNI-LOW","",DT_ZMDDL_ARM_DRIVER_0301_DELIVERY,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ActivateNodeGuiTree(0, "#1;#4")
Call ClickButton("btn\[25\]",False)
Call TakeScreenShot

Call GetTextboxValue("LIKP-VBELN", "", "DT_ZMDDL_ARM_DRIVER_0300_INBOUND_DELIVERY_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDDL_ARM_DRIVER_0300_INBOUND_DELIVERY_OUTPUT",DT_ZMDDL_ARM_DRIVER_0300_INBOUND_DELIVERY)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Post Goods Receipt   \(Shift\+F8\)",False)
Call TakeScreenShot

Call GetStatusBar("item2", "DT_ZMDDL_ARM_DRIVER_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
'Call VerifyStatusBar("Goods receipt for inbound delivery "&DT_ZMDDL_ARM_DRIVER_0300_INBOUND_DELIVERY_OUTPUT&" posted (article document "&DT_ZMDDL_ARM_DRIVER_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZMDDL_ARM_DRIVER_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ZMDDL_ARM_DRIVER_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(lcase(DT_ZMDDL_ARM_DRIVER_0300_CHECK_TEXT_OF_STATUSBAR))


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



