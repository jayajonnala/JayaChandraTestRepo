
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_ASS_01_09_009 - Create New planogram and assign Affiliates stores-check SAP
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

gstrTestCaseName = "Test_ASS_01_09_009 Affiliates stores-check SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_ASS_01_09_009 - Create New planogram and assign Affiliates store.xls"
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
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Assortment","V_WRS1-ASORT","",DT_WSOA1_0001_ASSORTMENT,False)
Call SetTextbox("Assortment Type","WRS1-ASSORTYP","",DT_WSOA1_0001_ASSORTMENT_TYPE,False)
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Assortment","V_WRS1-NAME1","",DT_WSOA1_0010_ASSORTMENT,False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Sales Organization","V_WRS1-VKORG","",DT_WSOA1_0015_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","V_WRS1-VTWEG","",DT_WSOA1_0015_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Listing procedure","V_WRS1-LSTFL","",DT_WSOA1_0015_LISTING_PROCEDURE,False)
Call SetTextbox("Layout Module","WRS1-LAYGR","",DT_WSOA1_0015_LAYOUT_MODULE,False)
Call SetTextbox("Format","WRS1-ASSORDIMVAL1","",DT_WSOA1_0015_FORMAT,False)
Call SetTextbox("Planogram Relevance","WRS1-ASSORDIMVAL2","",DT_WSOA1_0015_PLANOGRAM_RELEVANCE,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()

Call SelectTab("ASORT_TAB", "Langs", False)
Call SetTableData("WRFM_WSO6DYN020", "Lang.", 2, "", "", DT_WSOA1_0020_TABLECELL_LANG_1, false)
Call SetTableData("WRFM_WSO6DYN020", "Description", 2, "", "", DT_WSOA1_0020_TABLECELL_DESCRIPTION_1, false)
Call PressEnter()
Call TakeScreenShot

Call SelectTab("ASORT_TAB", "Assortment User", False)
Call SetTableData("WRFM_WSO6DYN040", "CustomerNoSite", 1, "", "", DT_WSOA1_0040_TABLECELL_CUSTOMERNOSITE_0, false)
Call SetTableData("WRFM_WSO6DYN040", "CustomerNoSite", 2, "", "", DT_WSOA1_0040_TABLECELL_CUSTOMERNOSITE_1, false)
Call SetTableData("WRFM_WSO6DYN040", "CustomerNoSite", 3, "", "", DT_WSOA1_0040_TABLECELL_CUSTOMERNOSITE_2, false)
Call SetTableData("WRFM_WSO6DYN040", "CustomerNoSite", 4, "", "", DT_WSOA1_0040_TABLECELL_CUSTOMERNOSITE_3, false)
Call SetTableData("WRFM_WSO6DYN040", "CustomerNoSite", 5, "", "", DT_WSOA1_0040_TABLECELL_CUSTOMERNOSITE_4, false)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Create Article Segments Using an Update Process   \(Ctrl\+F3\)",True)
Call TakeScreenShot
Call VerifyStatusBar(DT_WSOA1_0001_CHECK_TEXT_OF_STATUSBAR)

Call SetTcode(DT_WSOA1_0001_OKCD)     
Call PressEnter()     

Call SetTextbox("Assortment","V_WRS1-ASORT","",DT_WSOA1_0001_ASSORTMENT_OCC1,False)
Call PressEnter()
Call TakeScreenShot

Call VerifyTextBoxContent("Assortment","V_WRS1-ASORT","",DT_WSOA1_0010_CHECK_TEXT_OF_ASSORTMENT,False)
Call SelectTab("ASORT_TAB", "Assortment User", False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA1_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_0)
Call VerifyTableCellContent(2, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA1_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_1)
Call VerifyTableCellContent(3, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA1_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_2)
Call VerifyTableCellContent(4, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA1_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_3)
Call VerifyTableCellContent(5, "CustomerNoSite", "WRFM_WSO6DYN040", DT_WSOA1_0040_CHECK_TEXT_OF_TABLECELL_CUSTOMERNOSITE_4)
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




