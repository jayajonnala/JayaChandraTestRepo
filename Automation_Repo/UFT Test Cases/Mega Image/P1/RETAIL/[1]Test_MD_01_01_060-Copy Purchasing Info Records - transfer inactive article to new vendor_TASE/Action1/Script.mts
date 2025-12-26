
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor
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

gstrTestCaseName = "Test_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\TASE_DT_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectCheckbox("P_MRKVEN", 2, DT_ZMDPU_INFOREC_COPY_1000_MARK_NEW_VENDOR_AS_REGULAR, False)

Call SetTextbox("Purch. Organization","P_EKORG","",DT_ZMDPU_INFOREC_COPY_1000_PURCH_ORGANIZATION,false)
Call SetTextbox("Source Vendor","P_SRC_V","",DT_ZMDPU_INFOREC_COPY_1000_SOURCE_VENDOR,false)
Call SetTextbox("Target Vendor","P_TGT_V","",DT_ZMDPU_INFOREC_COPY_1000_P_TGT_V,false)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDPU_INFOREC_COPY_1000_ARTICLE,false)

Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call SelectRowGuiGrid("", 0, "Currency", "RON", false)
Call ClickButton("PROCESS   \(F9\)",False)
Call ClickButton("Yes",True)
Call TakeScreenShot

Call VerifyGridCellContentbyName("shell", 1, "Status", "", DT_ZMDPU_INFOREC_COPY_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATS)
Call ClickCellGuiGrid("", "", "Inforecord Number", 1, "", "", False)
Call TakeScreenShot

Call GetTextboxValue("EINA-INFNR", "", DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_INFO_RECORD_OUTPUT, False)

'''--------TransactionCode-SE16N----------''''

Call SetTcode(DT_ZMDPU_INFOREC_COPY_0101_OKCD)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Table","GD-TAB","",DT_ZMDPU_INFOREC_COPY_0100_TABLE,false)
Call PressEnter()

Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 2, "", "", DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_INFO_RECORD_OUTPUT, False)
Call TakeScreenShot
Call ClickButton("Online   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContentbyName("shell", 1, "Purchasing info rec.", "", DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_INFO_RECORD_OUTPUT)

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




