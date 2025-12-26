
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PM001 - Create Maintenance Notification for Echipamente frigorif 
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


gstrTestCaseName = "Test_PM001 - Create Maintenance Notification for Echipamente frigorif"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_PM001 - Create Maintenance Notification for Echipamente frigorif.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'''
'
'Call WriteRunTimeDataToExcelGlobalSheet("DT_INCREMENT",Cint(DT_INCREMENT)+1)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''--------TransactionCode-XK02----------''''


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectCheckbox("RF02K-D0110", 0, DT_XK02_0101_ADDRESS, False)
'Call SetTextbox("Supplier","RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
''Call SetTextbox("Vendor","RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Other E-Mail Addresses \(Others Exist\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(3, "Notes", "SAPLSZA6T_CONTROL6", DT_XK02_0600_CHECK_TEXT_OF_ORD_PM_RO)
Call SetTableData("SAPLSZA6T_CONTROL6", "E-Mail Address", 3, "Notes", DT_XK02_0600_CHECK_TEXT_OF_ORD_PM_RO, DT_XK02_0600_TABLECELL_EMAIL_ADDRESS_2, False)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(Enter\)",False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Continue \(Enter\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_XK02_0101_CHECK_TEXT_OF_SAVECHANGES)
Call PressEnter() 

'''''--------TransactionCode-/niw26----------''''

Call SetTcode(DT_XK02_0101_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call VerifyTextBoxContent("Notific. Status", "RIWO00-STTXT", "", DT_XK02_1050_CHECK_TEXT_OF_NOTIFIC_OSNO, False)
Call SetTextbox("Notification","VIQMEL-QMTXT","",DT_NUMBEROFNOTIFICATION,False)
Call SetTextbox("Functional loc.","RIWO1-TPLNR","",DT_XK02_0100_FUNCTIONAL_LOC,False)
Call SetTextbox("Equipment","RIWO1-EQUNR","",DT_XK02_0100_EQUIPMENT,False)
Call TakeScreenShot
Call SetTextbox("Coding","VIQMEL-QMCOD","",DT_XK02_7715_CODING,False)
Call TakeScreenShot
Call SetTextbox("Intervention Type","VIQMEL-ZZMATNR","",DT_XK02_0100_INTERVENTION_TYPE,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SelectTab("TAB_GROUP_10", "Dates", False)
Call SetTextbox("Notif.date","VIQMEL-QMDAT","",ConvertDate(DT_XK02_7316_NOTIFDATE),False)
Call SetTextbox("Required Start","VIQMEL-STRMN","",ConvertDate(DT_XK02_7316_REQUIRED_START),False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
'''''''***************Entering Mandatory partner**********************
'If SAPGuiSession("Session").SAPGuiWindow("Information").SAPGuiEdit("Information Message").GetROProperty("value") = DT_XK02_1050_INFORMATION Then
	Call TakeScreenShot
	Call ClickButtonIfExist("Continue   \(Enter\)",False)
	Call SetTableData("SAPLIPARTCTRL_0200", "Partner", 1, "Funct", "User Responsible", DT_XK02_1050_PARTNAME, False)
	Call TakeScreenShot
	Call ClickButtonIfExist("Back   \(F3\)",False)
'End If
Call GetStatusBar("item1", "DT_SAVENOTIF_OUTPUT")
Call VerifyStatusBar("Notification "&DT_SAVENOTIF_OUTPUT&" saved")

'''''--------TransactionCode-/niw28----------''''

Call SetTcode(DT_XK02_7200_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Notification","QMNUM-LOW","",DT_SAVENOTIF_OUTPUT,False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Notific. Status", "RIWO00-STTXT", "", DT_XK02_1050_CHECK_TEXT_OF_NOTIFIC_OSNO2, False)
Call VerifyTextBoxContent("Coding", "VIQMEL-QMGRP", "", DT_XK02_7715_CHECK_TEXT_OF_CODING, False)
Call VerifyTableCellContent(3, "Funct", "SAPLIPARTCTRL_0200", lcase(DT_XK02_0201_CHECK_VALUE_OF_VENDOR))
Call VerifyTableCellContent(3, "Partner", "SAPLIPARTCTRL_0200", DT_XK02_0201_CHECK_TEXT_OF_TABLECELL_PARTNER_2)
Call VerifyTableCellContent(3, "Name", "SAPLIPARTCTRL_0200", DT_XK02_0201_CHECK_TEXT_OF_NAME_2)
Call SelectTab("TAB_GROUP_10", "Dates", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Notif.date","VIQMEL-QMDAT","",ConvertDate(DT_XK02_7316_CHECK_TEXT_OF_NOTIFDATEFUTURE),False)
Call VerifyTextBoxContent("Required Start","VIQMEL-STRMN","",ConvertDate(DT_XK02_7316_CHECK_TEXT_OF_REQUIREDDATE_START),False)
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


