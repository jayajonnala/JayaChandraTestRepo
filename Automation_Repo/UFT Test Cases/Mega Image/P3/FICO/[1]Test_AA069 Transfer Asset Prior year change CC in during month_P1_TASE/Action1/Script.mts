
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA069 Transfer Asset Prior year change CC in during month_P1
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


gstrTestCaseName = "Test_AA069 Transfer Asset Prior year change CC in during month_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_AA069 Transfer Asset Prior year change CC in during month_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call EndDateof445PeriodByDate(DT_DATE,"DT_END_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()




''--------TransactionCode-AS02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''All selections (shift+F7)
Call ClickButtonIfExist("All Selections   \(Shift\+F7\)", False)
Call TakeScreenShot

'' SelectRadioButtonIfExist(radiobuttonName, radiobuttonIndex, blnIsItPopup)
Call SelectRadioButton("XEINZEL", "List assets", False)

Call SetTextbox("Company code","BUKRS-LOW", "", DT_S_ALR_87011966_1000_COMPANY_CODE, False)
Call SetTextbox("Asset number","ANLAGE-LOW", "", DT_S_ALR_87011966_1000_ASSET_NUMBER, False)
Call SetTextbox("Subnumber","UNTNR-LOW", "", DT_S_ALR_87011966_1000_SUBNUMBER, False)

Call SetTextbox("Cost center","SO_KOSTL-LOW", "", DT_S_ALR_87011966_1000_COST_CENTER, False)
Call SetTextbox("Sort Variant","SRTVR", "", DT_S_ALR_87011966_1000_SORT_VARIANT, False)
Call SetTextbox("Display variant","P_VARI", "", DT_S_ALR_87011966_1000_DISPLAY_VARIANT, False)
Call SetTextbox("Report date","BERDATUM", "", DT_REPORT_DATE, False)
 Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot
''Call VerifyTextPopup("Processing of batch input session completed")
If SAPGuiSession("type:=GuiSession").SAPGuiWindow("text:=Information").Exist(0)  Then
	SAPGuiSession("type:=GuiSession").SAPGuiWindow("text:=Information").Activate
	Call VerifyTextBoxContent("Information Message", "MESSTXT1", 0, lcase("No data was selected"), True)
End If
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot

Call SetTextbox("Cost center","SO_KOSTL-LOW", "", DT_S_ALR_87011966_1000_COST_CENTER_OCC1, False)
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "Asset", 0, DT_S_ALR_87011966_1000_ASSET_NUMBER)
Call VerifyGridCellContent("", 1, "Sub-number", 0, DT_S_ALR_87011966_1000_SUBNUMBER)

''''--------TransactionCode-/NEX ----------''''
''Call SetTcode(DT_S_ALR_87011966_0500_OKCD)
''Call PressEnter()     
''''Call TakeScreenShot
''


Call LogOff()

Call FinalStatus ()







'''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
'''Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'''Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'''Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 
''
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
''
'''// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''' ................NOTE: 
'''.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'''.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'''.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'''.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
''' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''

