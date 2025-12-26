

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_017_Creation_of_SPC_Centralized_article_and_margin_logic
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


gstrTestCaseName = "Test_01PRI00_017_Creation_of_SPC_Centralized_article_and_margin_logic"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_017_Creation_of_SPC_Centralized_article_and_margin_logic_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'
''--------------------------------------------  VKP5------------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
''grid title changed
''Call SelectRowGuiGrid("Variant Catalog for Program RWVKP007.*","","Variant name",DT_VKP5_VARIANT,True)
Call SelectRowGuiGrid("Variant Catalog.*","","Variant name",DT_VKP5_VARIANT,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_VKP5_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_VKP5_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call SelectCellGuiGrid("",0,1,"Vendor",False)
Call ClickButtonToolBar("KALK",0)
Call TakeScreenShot()
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",3,"","","DT_VKP5_6201_GET_TEXT_OF_TABLECELL_AMOUNT_2_OUTPUT",False)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",7,"","","DT_VKP5_6201_GET_TEXT_OF_TABLECELL_AMOUNT_6_OUTPUT",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call GetGridContent("",0,"EKPNN",1,"<NA>","<NA>","DT_VKP5_100_GETCELLVALUE_OF_GRIDCELL_0_EKPNN_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",1,"ENDPR",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)

Call SelectCellGuiGrid("",0,2,"Vendor",False)
Call ClickButtonToolBar("KALK",0)
Call TakeScreenShot()
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",3,"","","DT_VKP5_6201_GET_TEXT_OF_TABLECELL_AMOUNT_2_OCC2_OUTPUT",False)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",7,"","","DT_VKP5_6201_GET_TEXT_OF_TABLECELL_AMOUNT_6_OCC2_OUTPUT",False)
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Call GetGridContent("",0,"EKPNN",1,"<NA>","<NA>","DT_VKP5_100_GETCELLVALUE_OF_GRIDCELL_1_EKPNN_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",2,"ENDPR",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ENDPR)
Call SelectAllRowGuiGrid("",0,False)
Call  ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("item1","DT_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBarMessageType("S")

Call LogOff()
Call FinalStatus ()



