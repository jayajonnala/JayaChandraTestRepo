
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_ASS_01_09_011-Remove several articles from the planogram created in ASS_01_09_010 Check SAP
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

gstrTestCaseName = "Test_ASS_01_09_011- ASS_01_09_010 Check SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_ASS_01_09_011-Remove several articles from the planogram created in ASS_01_09_010  Check SAP.xls"
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
Call SelectTab("TABSTRIP_ENTRY", "Start: Layout Modules", False)
Call SetTextbox("Layout Module","SO_LAYGR-LOW","",DT_WLWBN_0020_LAYOUT_MODULE,False)
Call SetTextbox("Selection Date","G_DATE","",ConvertDate(DT_WLWBN_0020_SELECTION_DATE),False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call ActivateNodeGuiTree(0, "#1;#2;#2")
Call ClickButton("Change and Delete Article  Assignments   \(Shift\+F5\)",False)
Call SelectColumnGuiGrid("Selection of Article", 1, "MATNR", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call ClickButtonIfExist("%_%%DYN001_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_WLWBN_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_WLWBN_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButton("Execute   \(Enter\)",True)
Call SelectAllRowGuiGrid("Selection of Article", 1, False)
Call ClickButtonIfExist("Change Article Placing Details   \(Shift\+F8\)",False)
Call TakeScreenShot
Call SelectAllRowGuiGrid("Mainten. of Article Assignm.", 1, False)
Call ClickButtonIfExist("Delete All Article Assignments on Change Level   \(Ctrl\+Shift\+F2\)",False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_WLWBN_0400_CHECK_TEXT_OF_STATUSBAR)
Call ClickButtonIfExist("Back   \(F3\)",False)

''''''--------TransactionCode-/nWSM4L---------''''

Call SetTcode(DT_WLWBN_0400_OKCD)     
Call PressEnter()  

Call SetTextbox("Date From","DATUM_AB","",ConvertDate(DT_WLWBN_1000_DATE_FROM),False)
Call SetTextbox("Assortment","FILIA-LOW","",DT_WLWBN_1000_ASSORTMENT,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call ActivateNodeGuiTree(0, "#1")
Call ActivateNodeGuiTree(0, "#2")
Call TakeScreenShot

''''''--------TransactionCode-/nZMDAS_WSL11---------''''

Call SetTcode(DT_WLWBN_0100_OKCD)     
Call PressEnter()   

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_WLWBN_1000_ASSORTMENT_OCC1,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_WLWBN_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_WLWBN_1000_VALID_FROM),False)
Call SetTextbox("Valid To","P_DATBI","",ConvertDate(DT_WLWBN_1000_VALID_TO),False)
Call SetTextbox("Customer No. - Site","S_LOCNR-LOW","",DT_SITE,False)
Call SetTextbox("to","S_LOCNR-HIGH","",DT_SITE_TO,False)
Call FocusTextBox("Article", "S_ARTNR-LOW", False)
Call ClickButton("%_S_ARTNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_WLWBN_3010_TABLECELL_SINGLE_VALUE_1_OCC1, True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar(DT_WLWBN_1000_CHECK_TEXT_OF_STATUSBAR)


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




