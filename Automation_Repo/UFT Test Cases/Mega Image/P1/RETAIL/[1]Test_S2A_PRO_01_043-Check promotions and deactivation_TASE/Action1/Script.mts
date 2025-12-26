
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_043-Check promotions and deactivation 
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


gstrTestCaseName = "Test_S2A_PRO_01_043-Check deactivation"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''-----------------------Login-----------------------------'''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''''''-----------------------TCode WAK5-----------------------------'''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()
Call TakeScreenShot()


Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK5_1300_PROMOTION,False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call TakeScreenShot()
Call ClickButton("Select All",False)
Call TakeScreenShot()
Call CLickButton("Deactivate prices",False)
Call TakeScreenShot()
Call CLickButtonIfExist("Execute   \(Enter\)",True)
Call TakeScreenShot()
Call CLickButtonIfExist("Yes",True)

Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")

Call VerifyStatusBar("Promotion "&DT_DOCUMENT_NUMBER_OUTPUT&" saved")

''''''-----------------------TCode WAK5-----------------------------'''

Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1)  
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_DOCUMENT_NUMBER_OUTPUT,False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call VerifyTableCellContent(1, "Sales pr.", "SAPMWAKAPREISAK", DT_WAK5_2400_CHECK_TEXT_OF_TABLECELL_SALES_PR_0)
Call TakeScreenShot()


Call LogOff()

Call FinalStatus()













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




