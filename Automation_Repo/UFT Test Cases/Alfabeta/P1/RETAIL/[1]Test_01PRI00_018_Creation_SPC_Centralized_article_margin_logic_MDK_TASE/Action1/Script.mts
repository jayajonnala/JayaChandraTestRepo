

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_018_Creation_SPC_Centralized_article_margin_logic_MDK
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

gstrTestCaseName = "Test_01PRI00_018_Creation_SPC_Centralized_article_margin_logic_MDK"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_018_Creation_SPC_Centralized_article_margin_logic_MDK_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
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
''--------------------------------------------  MEK3------------------------------------------------
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_MEK3_100_CONDITION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Purch\. Organization","F001","",DT_MEK3_1000_PURCH_ORGANIZATION,False)
Call ClickButton("%_F002_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_MEK3_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1,"","","DT_MEK3_1993_GET_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT",False)
Call GetTableCellData("SAPMV13ATCTRL_FAST_ENTRY","Amount",2,"","","DT_MEK3_1993_GET_TEXT_OF_TABLECELL_AMOUNT_1_OUTPUT",False)

''--------------------------------------------  VKP5------------------------------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MEK3_1993_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RWVKP007.*","","Variant name",DT_VKP5_VARIANT,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_MEK3_3010_TABLECELL_SINGLE_VALUE_0_OCC2,True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_MEK3_3010_TABLECELL_SINGLE_VALUE_1_OCC2,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call SetTextbox("Sales price determination seq\.","P_VKERV","",DT_PP_DETERM,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call VerifyGridCellContent("",1,"LIFNR",0,"")
Call VerifyGridCellContent("",2,"LIFNR",0,"")
Call VerifyGridCellContent("",1,"MATNR",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",2,"MATNR",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MATNR)
Call VerifyGridCellContent("",1,"EKPNN",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKPNN)
Call VerifyGridCellContent("",1,"EKPNN",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKPNN)
Call VerifyGridCellContent("",2,"EKPNN",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_EKPNN)

Call SelectCellGuiGrid("",0,1,"Vendor",False)
Call ClickButtonToolBar("KALK",0)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",3,"","","DT_MEK3_6201_GET_TEXT_OF_TABLECELL_AMOUNT_2_OUTPUT",False)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",7,"","","DT_MEK3_6201_GET_TEXT_OF_TABLECELL_AMOUNT_6_OUTPUT",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",1,"ENDPR",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)

Call SelectCellGuiGrid("",0,2,"Vendor",False)
Call ClickButtonToolBar("KALK",0)
Call TakeScreenShot()
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",3,"","","DT_MEK3_6201_GET_TEXT_OF_TABLECELL_AMOUNT_2_OCC2_OUTPUT",False)
Call GetTableCellData("SAPLV69ATCTRL_KONDITIONEN","Amount",7,"","","DT_MEK3_6201_GET_TEXT_OF_TABLECELL_AMOUNT_6_OCC2_OUTPUT",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",2,"ENDPR",0,DT_MEK3_100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ENDPR)

Call SelectAllRowGuiGrid("",0,False)
Call  ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("item1","DT_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBarMessageType("S")


Call LogOff()
Call FinalStatus ()



