
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_MD_01_01_250 Select several articles Mntn all char
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
gstrTestCaseName = "Test_PRE_MD Mntn all char"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_Pre_MD_01_01_250 Select several articles Mntn all char.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE,false)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM43_0100_SALES_ORG,false)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM43_0100_DISTR_CHANNEL,false)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","","",false)
'''Call SetTextbox("Vendor","RMMW1-LIFNR","","",false)
Call SetTextboxNoLabel("RMMW1-LIFNR","","",false)

Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Next Page",False)
Call ClickButton("Next Page",False)
Call ClickButton("Next Page",False)
'Commented this line of code as the Textt box is availabl ein page 3 it self.
Call ClickButton("Next Page",False)
Call VerifyTextBoxContent("Electronic shelf label", "RCTMS-MWERT", "", "", False)
Call TakeScreenShot

Call ClickButton("Back   \(F3\)",False)

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM43_0100_ARTICLE_OCC1,false)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM43_0100_SALES_ORG_OCC1,false)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM43_0100_DISTR_CHANNEL_OCC1,false)
Call SetTextbox("Purchasing Org.","RMMW1-EKORG","","",false)
'''Call SetTextbox("Vendor","RMMW1-LIFNR","","",false)
Call SetTextboxNoLabel("RMMW1-LIFNR","","",false)
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Next Page",False)
Call ClickButton("Next Page",False)
Call VerifyTextBoxContent("Electronic shelf label", "RCTMS-MWERT", "", "", False)
Call TakeScreenShot

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




