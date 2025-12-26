
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P2
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P2
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_S2A_PRO_01_015- new tariff_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRO_01_015-Create low level promotion New Store opening new tariff_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()     

Call TakeScreenShot()


Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION,False)

Call PressEnter()

Call TakeScreenShot()

Call SetTableData("SAPMWAKAUEB_SERF", "Article", 1, "", "", DT_WAK2_8210_TABLECELL_ARTICLE_0, False)

Call SetTableData("SAPMWAKAUEB_SERF", "Sales Price", 1, "", "", DT_WAK2_8210_TABLECELL_SALES_PRICE_0, False)

Call PressEnter()

Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call TakeScreenShot()

Call GetStatusBar("text","DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar(DT_WAK2_1100_GET_TEXT_OF_STATUSBAR_OUTPUT)

'--------------------------------------------------------
Call SetTcode(DT_WAK2_1100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION,False)

Call PressEnter()

Call ClickButton("MARKALL",False)

Call ClickButton("POSDETAIL",False)

Call TakeScreenShot()
Call CLickButton("Price activation   \(F6\)",True)

Call TakeScreenShot()

Call SelectMenuBar("Promotion;Check")

Call VerifyifGuiLabelExists(DT_WAK2_0120_CHECK_TEXT_OF_PRODUCT_CATALOGS)

Call ClickLabel("4",0, False)

Call VerifyifGuiLabelExists(DT_WAK2_0120_CHECK_TEXT_OF_MT04)

Call ClickLabel("4",1,False)
'Call ClickLabel("5",0,False)

'Call VerifyifGuiLabelExists(DT_WAK2_0120_CHECK_TEXT_OF_MI_PLAZA)
' 'VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
Call VerifyifGuiLabelExistsByRelativeid(DT_WAK2_0120_CHECK_TEXT_OF_MI_PLAZA,"wnd\[0\]/usr/lbl\[12,11\]")

Call ClickLabel("4",2, False)

Call VerifyifGuiLabelExists(DT_CHECK_ARTICLE_1)

Call ClickButton("Back   \(F3\)",False)

''Call ClickButton("btn\[11\]",False)
''Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Yes",True)

Call GetStatusBar("item1","DT_WAK2_1300_GET_TEXT_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Promotion "&DT_WAK2_1300_GET_TEXT_OF_STATUSBAR_OUTPUT&" saved")
'Call VerifyStatusBar(DT_WAK2_1300_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call LogOff()
Call FinalStatus()










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




