
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_ASS_01_09_010-Create New planogram. Assign multiple articles and stores_P2_Add articles
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

gstrTestCaseName = "Test_ASS_01_09_010_P2_Add articles"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_ASS_01_09_010-Create New planogram. Assign multiple articles and stores_P2.xls"
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
Call SetTextbox("Assortment","S_FILIA-LOW","",DT_ZMDAS_WSL11_1000_ASSORTMENT,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_ZMDAS_WSL11_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_ZMDAS_WSL11_1000_VALID_FROM),False)
Call SetTextbox("Valid To","P_DATBI","",ConvertDate(DT_ZMDAS_WSL11_1000_VALID_TO),False)
Call SetTextbox("Customer No. - Site","S_LOCNR-LOW","",DT_SITE,False)
Call SetTextbox("to","S_LOCNR-HIGH","",DT_SITE_TO,False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar(DT_ZMDAS_WSL11_1000_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode-/nWLWBN----------''''

Call SetTcode(DT_ZMDAS_WSL11_1000_OKCD)     
Call PressEnter()     

Call SelectTab("TABSTRIP_ENTRY", "Start: Layout Modules", False)
Call SetTextbox("Layout Module","SO_LAYGR-LOW","",DT_ZMDAS_WSL11_0020_LAYOUT_MODULE,False)
Call SetTextbox("Selection Date","G_DATE","",ConvertDate(DT_ZMDAS_WSL11_0020_SELECTION_DATE),False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call ActivateNodeGuiTree(0, "#1;#2;#2")
Call ClickButton("Assign Articles   \(Shift\+F2\)",False)
Call ClickButtonIfExist("%_SO_MATNR_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_ZMDAS_WSL11_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_ZMDAS_WSL11_3010_TABLECELL_SINGLE_VALUE_1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_ZMDAS_WSL11_3010_TABLECELL_SINGLE_VALUE_2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_ZMDAS_WSL11_3010_TABLECELL_SINGLE_VALUE_3, True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SelectColumnGuiGrid("Selection of Article", 1, "Unit of Measure", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call ClickButtonIfExist("%_%%DYN001_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_ZMDAS_WSL11_3010_TABLECELL_SINGLE_VALUE_0_OCC1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_ZMDAS_WSL11_3010_TABLECELL_SINGLE_VALUE_1_OCC1, True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call TakeScreenShot
Call SelectAllRowGuiGrid("Selection of Article", 1, False)
Call ClickButtonIfExist("Assign Selected Articles   \(Shift\+F1\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_ZMDAS_WSL11_0400_CHECK_TEXT_OF_STATUSBAR)


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




