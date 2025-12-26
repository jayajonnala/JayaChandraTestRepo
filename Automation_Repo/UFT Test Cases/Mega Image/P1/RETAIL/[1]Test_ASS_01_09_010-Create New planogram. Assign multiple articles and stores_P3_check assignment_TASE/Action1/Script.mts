
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_ASS_01_09_010-Create New planogram. Assign multiple articles and stores_P3_check assignment
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

gstrTestCaseName = "Test_ASS_01_09_010_P3_check assignment"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_ASS_01_09_010-Create New planogram. Assign multiple articles and stores_P3.xls"
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
Call SetTextbox("Date From","DATUM_AB","",ConvertDate(DT_WSM4L_1000_DATE_FROM),False)
Call SetTextbox("Assortment","FILIA-LOW","",DT_WSM4L_1000_ASSORTMENT,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call ActivateNodeGuiTree(0, "#1")
Call ActivateNodeGuiTree(0, "#2")
Call TakeScreenShot

''''''--------TransactionCode-/nZMDAS_WSL11---------''''

Call SetTcode(DT_WSM4L_0100_OKCD)     
Call PressEnter()   

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_WSM4L_1000_ASSORTMENT_OCC1,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_WSM4L_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_WSM4L_1000_VALID_FROM),False)
Call SetTextbox("Valid To","P_DATBI","",ConvertDate(DT_WSM4L_1000_VALID_TO),False)
Call SetTextbox("Customer No. - Site","S_LOCNR-LOW","",DT_SITE,False)
Call SetTextbox("to","S_LOCNR-HIGH","",DT_SITE_TO,False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectColumnGuiGrid("", 0, "Article", False)
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call SetTextbox("Article","%%DYN001-LOW","",DT_WSM4L_1105_ARTICLE,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "ARTNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR)
Call VerifyGridCellContent("", 1, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR)
Call VerifyGridCellContent("", 8, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_LOCNR)

Call SelectColumnGuiGrid("", 0, "Article", False)
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call SetTextbox("Article","%%DYN001-LOW","",DT_WSM4L_1105_ARTICLE_OCC1,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "ARTNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR_OCC1)
Call VerifyGridCellContent("", 1, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR_OCC1)
Call VerifyGridCellContent("", 8, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_LOCNR_OCC1)

Call SelectColumnGuiGrid("", 0, "Article", False)
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call SetTextbox("Article","%%DYN001-LOW","",DT_WSM4L_1105_ARTICLE_OCC2,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "ARTNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR_OCC2)
Call VerifyGridCellContent("", 1, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR_OCC2)
Call VerifyGridCellContent("", 8, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_LOCNR_OCC2)

Call SelectColumnGuiGrid("", 0, "Article", False)
Call ClickButton("Set filter   \(Ctrl\+F5\)",False)
Call SetTextbox("Article","%%DYN001-LOW","",DT_WSM4L_1105_ARTICLE_OCC3,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "ARTNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR_OCC3)
Call VerifyGridCellContent("", 1, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOCNR_OCC3)
Call VerifyGridCellContent("", 8, "LOCNR", 0, DT_WSM4L_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_LOCNR_OCC3)



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




