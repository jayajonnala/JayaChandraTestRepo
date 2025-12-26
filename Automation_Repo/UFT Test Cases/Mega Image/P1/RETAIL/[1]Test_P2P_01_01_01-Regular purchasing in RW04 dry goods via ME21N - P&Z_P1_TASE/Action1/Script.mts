
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_01-Regular purchasing in RW04 dry goods via ME21N - P&Z_P1
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


gstrTestCaseName = "Test_P2P_01_01_01- RW04 dry goods via ME21N - P&Z_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''---------Login-------------''''''
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''''-------Transaction Code ME21N-------'''''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)

Call VerifyTableCellContent(1, "Description", "SAPDV70ATC_NAST3", lcase(DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_0))
Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(4, "Description", "SAPDV70ATC_NAST3", lcase(DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_3))
Call VerifyTableCellContent(4, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_3)
Call TakeScreenShot

''''''-------Transaction Code /nME9F-------'''''''

Call SetTcode(DT_ME23N_0100_OKCD)     
Call PressEnter()     

Call SetTextbox("Application","P_KAPPL","",DT_ME23N_1000_APPLICATION,false)
Call SetTextbox("Processing Status","P_VSTAT","",DT_ME23N_1000_PROCESSING_STATUS,false)
Call SetTextbox("Document Number","S_EBELN-LOW","",DT_ME23N_1000_DOCUMENT_NUMBER,false)
Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME23N_1000_PURCHASING_ORGANIZATION,false)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectRowGuiGrid("", 0, "Message type", "ZNEU", False)
Call ClickButton("Display Message   \(Shift\+F8\)",False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call SelectRowGuiGrid("", 0, "Message type", "ZNEU", False)
Call ClickButton("Output Message   \(Shift\+F5\)",False)

''''''-------Transaction Code /nme23n-------'''''''

Call SetTcode(DT_ME23N_0500_OKCD)     
Call PressEnter() 

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER_OCC1,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call VerifyTableCellContent(3, "Description", "SAPDV70ATC_NAST3", lcase(DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_2))
Call VerifyTableCellContent(3, "Status", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_2)
Call LogOff()

Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


