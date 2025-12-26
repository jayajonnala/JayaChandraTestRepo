
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_176-Perform Initial Listing and check distr. channel
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

gstrTestCaseName = "Test_MD_01_01_176-Perform Initial Listing and check distr. channel"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_MD_01_01_176-Perform Initial Listing and check distr. channel.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-WSM3----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
Call ClickContextButtonToolBar("&FIND", 0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_WSM3_0841_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
''The grid title has been changed hence, making them empty
'Call SelectRowGuiGrid("Variant Catalog for Program RWSORLIARTI","","Variant name",DT_WSM3_0841_SEARCH_TERM,True)
Call SelectRowGuiGrid("Variant Catalog.*","","Variant name",DT_WSM3_0841_SEARCH_TERM,True)
Call ClickButtonIfExist("Choose \(F2\)",True)
Call PressEnter()
Call SetTextbox("Article","MATNR-LOW","",DT_WSM3_1000_ARTICLE,False)
Call ClickButton("Execute   \(F8\)",False)

Call VerifyStatusBar(DT_WSM3_0120_CHECK_TEXT_OF_STATUSBAR)

''''--------TransactionCode-SE16N----------''''

Call SetTcode(DT_WSM3_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Table","GD-TAB","",DT_WSM3_0100_TABLE,false)
Call PressEnter()
Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 2, "", "", DT_WSM3_0100_TABLECELL_FRVALUE_1, False)
Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 3, "", "", DT_WSM3_0100_TABLECELL_SALESORG_1, False)
Call SetTableData("SAPLSE16NSELFIELDS_TC", "To value", 3, "", "", DT_WSM3_0100_TABLECELL_SALESORG_2, False)
Call TakeScreenShot
Call ClickButton("Online   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell", 1, "Article", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContentbyName("shell", 1, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG)
Call VerifyGridCellContentbyName("shell", 2, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_VTWEG)
Call VerifyGridCellContentbyName("shell", 3, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VTWEG)
Call VerifyGridCellContentbyName("shell", 4, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VTWEG)
Call VerifyGridCellContentbyName("shell", 5, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_4_VTWEG)
Call VerifyGridCellContentbyName("shell", 6, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_5_VTWEG)
Call VerifyGridCellContentbyName("shell", 7, "Distribution Channel", "", DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_6_VTWEG)


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




