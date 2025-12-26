

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_026_Creation_of_Sales_price_calulation_for_articles 
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

gstrTestCaseName = "Test_01PRI00_026_calulation_for_articles"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_026_Creation_of_Sales_price_calulation_for_articles_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

''--------------------------------------------  VKP5------------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
'''grid title changed
''Call SelectRowGuiGrid("Variant Catalog for Program RWVKP007.*","","Variant name",DT_VKP5_VARIANT,True)
Call SelectRowGuiGrid("Variant Catalog.*","","Variant name",DT_VKP5_VARIANT,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call VerifyGridCellContent("",1,"LIFNR",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)	
Call VerifyGridCellContent("",1,"MATNR",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)	
Call VerifyGridCellContent("",1,"EKPNN",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKPNN)	
Call VerifyGridCellContent("",1,"SPANE",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SPANE)	
Call VerifyGridCellContent("",1,"ENDPR",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)	

Call SetGridData("",1,"ENDPR",DT_VKP5_100_GRIDCELL_0_FINPR,False)
Call PressEnter()  ' 
Call TakeScreenShot()

Call SelectAllRowGuiGrid("",0,False)
Call  ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("item1","DT_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBarMessageType("S")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_VKP5_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_VKP5_100_CONDITION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()

Call SelectRadioButton("RV130-SELKZ","Article per SOrg/DstCh",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call TakeScreenShot()

Call SetTextbox("Sales Organization","F001","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Valid On","SEL_DATE","",ConvertDate(DT_VKP5_1000_VALID_ON),False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call SelectRowGuiTable("SAPMV13ATCTRL_FAST_ENTRY","Article",DT_VKP5_1000_ARTICLE,False)
Call ClickButton("Condition Supplement   \(Shift\+F7\)",False)
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_D0201",DT_VKP5_201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)


Call LogOff()
Call FinalStatus ()



