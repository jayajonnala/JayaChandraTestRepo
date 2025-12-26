
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_090-Change EAN from article A to article B_p2
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

gstrTestCaseName = "Test_MD_01_01_090 A to article B_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_MD_01_01_090-Change EAN from article A to article B_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()



''--------TransactionCode-SE16N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


 Call SetTextbox("Table","GD-TAB", "", DT_SE16N_0100_TABLE, False)
 Call TakeScreenShot
 Call PressEnter() 

Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
Call SetTextbox("Field Name","SVALD-VALUE", "", "MATNR", True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call FindRowNumber("SAPLSE16NSELFIELDS_TC", "Technical name", "MATNR", "FIND_ROW_NUMBER_OUTPUT")

Call ClickCellTable("SAPLSE16NSELFIELDS_TC","More",FIND_ROW_NUMBER_OUTPUT,"Technical name","MATNR", False)


Call SetTableDataNoRef("SAPLSE16NMULTI_TC", "Fr.Value", 1, DT_SE16N_0001_TABLECELL_FRVALUE_0, True)
Call SetTableDataNoRef("SAPLSE16NMULTI_TC", "Fr.Value", 2, DT_SE16N_0001_TABLECELL_FRVALUE_1, True)
Call TakeScreenShot

Call ClickButtonIfExist("Transfer Data   \(F8\)", True)
Call TakeScreenShot

Call ClickButtonIfExist("Online   \(F8\)", False)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "Article", 0, DT_SE16N_0001_TABLECELL_FRVALUE_0)
Call VerifyGridCellContent("", 1, "GTIN", 0, DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EAN11)
Call VerifyGridCellContent("", 2, "Article", 0, DT_SE16N_0001_TABLECELL_FRVALUE_1)
Call VerifyGridCellContent("", 2, "GTIN", 0, DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_EAN11)

Call LogOff()

Call FinalStatus ()









'''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
'''Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'''Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'''Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 
''
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
''
'''// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''' ................NOTE: 
'''.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'''.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'''.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'''.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
''' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''

