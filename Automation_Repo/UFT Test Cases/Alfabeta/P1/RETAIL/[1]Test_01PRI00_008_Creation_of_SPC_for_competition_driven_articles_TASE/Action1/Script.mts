

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_008_Creation_of_SPC_for_competition_driven_articles
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


gstrTestCaseName = "Test_01PRI00_008_Creation_of_SPC_for_competition_driven_articles"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_008_Creation_of_SPC_for_competition_driven_articles_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =6
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
''Grid title has been changed
''Call SelectRowGuiGrid("Variant Catalog for Program RWVKP007.*","","Variant name",DT_VKP5_VARIANT,True)
Call SelectRowGuiGrid("Variant Catalog.*","","Variant name",DT_VKP5_VARIANT,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)
Call SetTextbox("Vendor \(External Procurement\)","P_LIFNR","","",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SelectMenuBar("Goto;Competitor Prices")
Call TakeScreenShot()
Call VerifyWindowValue(": Competitor Prices")
Call GetTextboxValue("\*CALP-VORPR","","DT_VKP5_200_GET_TEXT_OF_CALPVORPR_OUTPUT",False)
Call GetTextboxValue("\*CALP-ENDPR","","DT_VKP5_200_GET_TEXT_OF_CALPENDPR_OUTPUT",False)
Call ClickButton("Back   \(F3\)",False)

Call VerifyGridCellContent("",1,"VORPR",0,Convert4digitNumber(DT_VKP5_200_GET_TEXT_OF_CALPVORPR_OUTPUT))
Call GetGridContent("",0,"Sales Pr. from Cond.",1,"<NA>","<NA>","DT_VKP5_OLDPRICE_OUTPUT")
DT_OLDPRICE= DT_VKP5_OLDPRICE_OUTPUT
Call SetGridData("",1,"ENDPR",DT_VKP5_OLDPRICE_OUTPUT,False)
Call PressEnter() 

Call  ClickButton("btn\[11\]",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)
Call GetStatusBar("item1","DT_PRICING_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Data saved; pricing document "& DT_PRICING_DOCUMENT_OUTPUT &" created" )


'
''--------------------------------------------  VKP5------------------------------------------------
Call SetTcode(DT_VKP5_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Pricing document","P_KBELN","",DT_PRICING_DOCUMENT,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call SelectAllRowGuiGrid("",0,False)
Call ClickButtonToolBar("&SORT_DSC",0)
Call VerifyGridCellContent("",1,"MATNR",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",1,"ENDPR",0,DT_OLDPRICE)
Call VerifyGridCellContent("",1,"BPSTA",0,DT_VKP5_100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BPSTA)




Call LogOff()
Call FinalStatus ()



