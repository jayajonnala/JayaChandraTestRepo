
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test2_MI_Test_P2P_01_01_054-EDI process in RW04 for good receipt_P1_PO_Message
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


gstrTestCaseName = "Test_MD_01_01_088-Change dimensions-Weight for Base UoM"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_MD_01_01_088-Change dimensions-Weight for Base UoM.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-MM42----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,false)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,false)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,false)
Call PressEnter()
Call TakeScreenShot

Call GetTableCellData("SAPLMGD2TC_ME_8022", "Gross Weight", 1, "", "","DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_GROSS_WEIGHT_0_OUTPUT", False)
Call GetTableCellData("SAPLMGD2TC_ME_8022", "Net Weight", 1, "", "","DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_NET_WEIGHT_0_OUTPUT", False)
Call GetTableCellData("SAPLMGD2TC_ME_8022", "Length", 1, "", "","DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_LENGTH_0_OUTPUT", False)
Call GetTableCellData("SAPLMGD2TC_ME_8022", "Width", 1, "", "","DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_WIDTH_0_OUTPUT", False)
Call GetTableCellData("SAPLMGD2TC_ME_8022", "Height", 1, "", "","DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_HEIGHT_0_OUTPUT", False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableData("SAPLMGD2TC_ME_8022", "Gross Weight", 1, "", "", DT_MM42_8022_TABLECELL_GROSS_WEIGHT_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Net Weight", 1, "", "", DT_MM42_8022_TABLECELL_NET_WEIGHT_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Length", 1, "", "", DT_MM42_8022_TABLECELL_LENGTH_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Width", 1, "", "", DT_MM42_8022_TABLECELL_WIDTH_0, False)
Call SetTableData("SAPLMGD2TC_ME_8022", "Height", 1, "", "", DT_MM42_8022_TABLECELL_HEIGHT_0, False)

Call TakeScreenShot
Call PressEnter()

' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)


Call ClickButtonIfExist("Yes",True)

Call VerifyStatusBar("Article "&DT_CHECK_STATUS&" changed")


'''--------TransactionCode-MM43----------''''

Call SetTcode(DT_TRANSACTION_CODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE1,false)
Call SetTextbox("Sales Org.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG1,false)
Call SetTextbox("Distr. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL1,false)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100", 1, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call VerifyTableCellContent(1, "Gross Weight", "SAPLMGD2TC_ME_8022", DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_GROSS_WEIGHT_0_OCC1)
Call VerifyTableCellContent(1, "Net Weight", "SAPLMGD2TC_ME_8022", DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_NET_WEIGHT_0_OCC1)
Call VerifyTableCellContent(1, "Length", "SAPLMGD2TC_ME_8022", DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_LENGTH_0_OCC1)
Call VerifyTableCellContent(1, "Width", "SAPLMGD2TC_ME_8022", DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_WIDTH_0_OCC1)
Call VerifyTableCellContent(1, "Height", "SAPLMGD2TC_ME_8022", DT_MM42_8022_CHECK_TEXT_OF_TABLECELL_HEIGHT_0_OCC1)


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




