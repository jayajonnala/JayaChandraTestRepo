
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_079-List article in WH RW21
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
gstrTestCaseName = "Test_MD_01_01_079-List article in WH RW21"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_MD_01_01_079-List article in WH RW21.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''--------TransactionCode-WSM3----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
''grid title has been changed has regex is used
''Call SelectRowGuiGrid("Variant Catalog for Program RWSORLIARTI","","Variant name",DT_GRIDCELL_VARIANT_NAME,True)
Call SelectRowGuiGrid("Variant Catalog.*","","Variant name",DT_GRIDCELL_VARIANT_NAME,True)
Call ClickButtonIfExist("Choose \(F2\)",True)
Call PressEnter()
Call SetTextbox("Article","MATNR-LOW","",DT_WSM3_1000_ARTICLE,False)
Call ClickButton("Execute   \(F8\)",False)
Call VerifyifGuiLabelExists(DT_WSM3_0120_CHECK_TEXT_OF_START_DATE_23052018)
Call VerifyifGuiLabelExists(DT_WSM3_0120_CHECK_TEXT_OF_RW04)
Call VerifyifGuiLabelExistsByRelativeid(DT_WSM3_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[1,9\]")
Call VerifyifGuiLabelExists(DT_WSM3_0120_CHECK_TEXT_OF_RW21)
Call VerifyifGuiLabelExistsByRelativeid(DT_WSM3_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[1,12\]")
Call VerifyStatusBar(DT_WSM3_0120_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

''''''--------TransactionCode-SE16N----------''''

Call SetTcode(DT_WSM3_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Table","GD-TAB","",DT_WSM3_0100_TABLE,false)
Call PressEnter()
Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 3, "", "", DT_WSM3_0100_TABLECELL_FRVALUE_2, False)
Call ClickButton("Online   \(F8\)",False)
Call SelectColumnGuiGrid("", 0, "Assortment", False)
Call ClickButtonToolBar("&MB_FILTER",0)
Call ClickButtonIfExist("Multiple selection",0)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_WSM3_1105_ASSORTMENT, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_WSM3_3010_TABLECELL_SINGLE_VALUE_1, True)
Call ClickButton("Copy   \(F8\)",False)
Call ClickButtonIfExist("Execute   \(Enter\)", True)
Call VerifyGridCellContentbyName("shell",1,"Assortment","",DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FILIA)
Call VerifyGridCellContentbyName("shell",2,"Assortment","",DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_FILIA)
Call VerifyGridCellContentbyName("shell",1,"Article","",DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR)
Call VerifyGridCellContentbyName("shell",2,"Article","",DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ARTNR)
'Call VerifyGridCellContentbyName("shell",1,"Last Changed","",ConvertDate(DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAE))
'Call VerifyGridCellContentbyName("shell",2,"Last Changed","",ConvertDate(DT_WSM3_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAE))
Call TakeScreenShot

Call LogOff

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





