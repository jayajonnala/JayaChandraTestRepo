
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_ASS_01_09_010-Create New planogram. Assign multiple articles and stores_Add stores
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

gstrTestCaseName = "Test_ASS_01_09_010- articles and stores_Add stores"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_ASS_01_09_010-Create New planogram. Assign multiple articles and stores.xls"
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
Call SetTextbox("Layout Module","S_LAYGR-LOW","",DT_WSOA6_1110_LAYOUT_MODULE,False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#1;#1")
Call TakeScreenShot
Call ClickButtonToolBar("&FIND", 1)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_WSOA6_0840_CELL_CONTENT,True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonToolBar("ASSIGN", 1)
Call SetCombo("WRS1KEY-ASORT", DT_WSOA6_0700_ASSORTMENT)
Call ClickButtonIfExist("Assign Assortments   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonToolBar("&FIND", 1)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_WSOA6_0841_CELL_CONTENT,True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonToolBar("ASSIGN", 1)
Call SetCombo("WRS1KEY-ASORT", DT_WSOA6_0700_ASSORTMENT_OCC1)
Call ClickButtonIfExist("Assign Assortments   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonToolBar("&FIND", 1)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_WSOA6_0842_CELL_CONTENT,True)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonToolBar("ASSIGN", 1)
Call SetCombo("WRS1KEY-ASORT", DT_WSOA6_0700_ASSORTMENT_OCC2)
Call ClickButtonIfExist("Assign Assortments   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Create Article Segments Using an Update Process   \(Ctrl\+F3\)",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_WSOA6_0100_CHECK_TEXT_OF_STATUSBAR)

 
''''''--------TransactionCode-/nWSOA3----------''''

Call SetTcode(DT_WSOA6_0100_OKCD)     
Call PressEnter()     

Call SetTextbox("Assortment","V_WRS1-ASORT","",DT_WSOA6_0001_ASSORTMENT,False)
Call PressEnter()
Call TakeScreenShot

Call SelectTab("ASORT_TAB", "Assortment User", False)
Call TakeScreenShot
Call VerifyTableCellContent(6, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA6_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_5)
Call VerifyTableCellContent(7, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA6_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_6)
Call VerifyTableCellContent(8, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA6_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_7)


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




