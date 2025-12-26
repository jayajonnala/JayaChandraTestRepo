
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_06_01_064-Article to article transfers - same material type     
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


gstrTestCaseName = "Test_S2C_06_01_064- same material type"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2C_06_01_064-Article to article transfers - same material type-.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
'Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call PressEnter() 
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_R10)

Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("BUTTON_DETAIL",False)
Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE,False)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFME","",DT_MIGO_0390_QTY_IN_UNE_OCC1,False)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE,False)

'Call SetTextbox("Article Trfr Pstg","GOITEM-UMMAKTX","",DT_MIGO_0390_GOITEMUMMAKTX,False)
Call SetTextboxNoLabel("GOITEM-UMMAKTX","",DT_MIGO_0390_GOITEMUMMAKTX,False)
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE,False)
''''Call SetTextbox("Site Trfr Pstg","GOITEM-UMNAME1","",DT_MIGO_0390_GOITEMUMNAME1,False)
Call SetTextboxNoLabel("GOITEM-UMNAME1","",DT_MIGO_0390_GOITEMUMNAME1,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC,False)
''''Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIGO_0390_GOITEMUMLGOBE,False)
Call SetTextboxNoLabel("GOITEM-UMLGOBE","",DT_MIGO_0390_GOITEMUMLGOBE,False)
Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR&" posted")
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)
Call PressEnter()
Call TakeScreenShot
wait 2
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS_GOHEAD","Doc. info", False)
Call ClickButtonIfExist("FI Documents",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTON)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTON)


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


