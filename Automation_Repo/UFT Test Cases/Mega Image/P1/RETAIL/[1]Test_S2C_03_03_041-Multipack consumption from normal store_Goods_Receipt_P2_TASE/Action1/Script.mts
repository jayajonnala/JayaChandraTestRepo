
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_03_03_041-Multipack consumption from normal store_Goods_Receipt_P2     

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

gstrTestCaseName = "Test_S2C_03_03_041- store_Goods_Receipt_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2C_03_03_041-Multipack consumption from normal store_Goods_Receipt_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call PressEnter()
Call TakeScreenShot
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Close Detail Data",false)

Call SetTextBoxNoLabel("GODEFAULT_TV-BWART","0",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call PressEnter()
Call TakeScreenShot

''Call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,False)
''Call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",2,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_1,False)
''Call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",3,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_2,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",2,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",3,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_2,False)

Call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",2,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",3,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_2,False)

Call SetTableData("SAPLMIGOTV_GOITEM","EUn",1,"","",DT_MIGO_0200_TABLECELL_EUN_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","EUn",2,"","",DT_MIGO_0200_TABLECELL_EUN_1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","EUn",3,"","",DT_MIGO_0200_TABLECELL_EUN_2,False)

Call SetTableData("SAPLMIGOTV_GOITEM","SLoc",1,"","",DT_MIGO_0200_TABLECELL_SLOC_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","SLoc",2,"","",DT_MIGO_0200_TABLECELL_SLOC_1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","SLoc",3,"","",DT_MIGO_0200_TABLECELL_SLOC_2,False)

Call SetTableData("SAPLMIGOTV_GOITEM","#13",1,"","",DT_MIGO_0200_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","#13",2,"","",DT_MIGO_0200_TABLECELL_SITE_1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","#13",3,"","",DT_MIGO_0200_TABLECELL_SITE_2,False)
Call PressEnter()  
Call TakeScreenShot
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)
Call PressEnter()
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call PressEnter()
Call SetTextboxNolabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART_OCC1,False)
Call TakeScreenShot
'Call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0_OCC1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0_OCC1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_0_OCC1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","EUn",1,"","",DT_MIGO_0200_TABLECELL_EUN_0_OCC1,False)
Call SetTableData("SAPLMIGOTV_GOITEM","SLoc",1,"","",DT_MIGO_0200_TABLECELL_SLOC_0_OCC1,False)

'''Call SetTableData("SAPLMIGOTV_GOITEM","Site",1,"","",DT_MIGO_0200_TABLECELL_CUSTOMER_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Plnt",1,"","",DT_MIGO_0200_TABLECELL_CUSTOMER_0,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Cost Center",1,"","",DT_MIGO_0200_TABLECELL_BATCH_0,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" posted")

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


