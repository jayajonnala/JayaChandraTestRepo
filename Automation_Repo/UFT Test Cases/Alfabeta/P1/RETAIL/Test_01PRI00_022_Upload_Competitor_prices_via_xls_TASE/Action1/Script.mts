

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_022_Upload_Competitor_prices_via_xls
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

gstrTestCaseName = "Test_01PRI00_022_Upload_Competitor_prices_via_xls"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_022_Upload_Competitor_prices_via_xls_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'
''--------------------------------------------  ZMDPC_UPLOAD_COND------------------------------------------------
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program ZMDPC_UPLOAD_PRICES_FROM_FILE.*","","Variant name",DT_ZMDPC_UPLOAD_COND_600_GRIDCELL_3_VARIANT_NAME,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call SetTextbox("File name","P_FILE","",DT_ZMDPC_UPLOAD_COND_1000_FILE_NAME_OCC2,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call VerifyGridCellContent("",1,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT)
Call VerifyGridCellContent("",2,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LIGHT)
Call VerifyGridCellContent("",3,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LIGHT)
Call VerifyGridCellContent("",4,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_LIGHT)
Call VerifyGridCellContent("",4,"Message Text",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MSGTX)

Call GetGridContent("",0,"MATNR",1,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")
Call GetGridContent("",0,"KUNNR",1,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_0_KUNNR_OUTPUT")
Call GetGridContent("",0,"KUNNR",2,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_1_KUNNR_OUTPUT")
Call GetGridContent("",0,"KUNNR",3,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_2_KUNNR_OUTPUT")
Call GetGridContent("",0,"KBETR_EDIT",3,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GET_GRIDCELL_0_PRICEEDIT_OUTPUT")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''Call SetGridData("",1,"KBETR_EDIT",DT_ZMDPC_UPLOAD_COND_500_GRIDCELL_0_PRICEEDIT,False)
Call GetGridContent("",0,"KBETR_EDIT",1,"<NA>","<NA>","DT_GETCELLVALUE_OF_GRIDCELL_0_KBETR_EDIT_OUTPUT")
Call GetGridContent("",0,"KBETR_EDIT",2,"<NA>","<NA>","DT_GETCELLVALUE_OF_GRIDCELL_1_KBETR_EDIT_OUTPUT")
Call GetGridContent("",0,"KBETR_EDIT",3,"<NA>","<NA>","DT_GETCELLVALUE_OF_GRIDCELL_2_KBETR_EDIT_OUTPUT")

Call  ClickButton("Select All   \(F5\)",False)
Call TakeScreenShot()
Call  ClickButton("Create Conditions   \(F8\)",False)

Call VerifyGridCellContent("",1,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GRIDCELL_0_LIGHT_OCC2)
Call VerifyGridCellContent("",2,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GRIDCELL_1_LIGHT_OCC2)
Call VerifyGridCellContent("",3,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GRIDCELL_2_LIGHT_OCC2)
Call VerifyGridCellContent("",4,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GRIDCELL_3_LIGHT_OCC2)
Call  ClickButton("Deselect All   \(F6\)",False)
Call TakeScreenShot()
'
'--------------------------------------------  VK13------------------------------------------------
Call SetTcode(DT_ZMDPC_UPLOAD_COND_500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_ZMDPC_UPLOAD_COND_100_CONDITION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()     

Call SelectRadioButton("RV130-SELKZ","Sales Org\./Distribution Channel/Material/Sales Unit/Customer",True)
Call TakeScreenShot()
Call ClickButton("Choose   \(Enter\)",True)

Call SetTextbox("Sales Organization","F001","",DT_ZMDPC_UPLOAD_COND_1000_SALES_ORGANIZATION_OCC2,False)
Call SetTextbox("Distribution Channel","F002","",DT_ZMDPC_UPLOAD_COND_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Article","F003-LOW","",DT_ZMDPC_UPLOAD_COND_1000_MATERIAL_OCC2,False)
Call TakeScreenShot()

Call  ClickButton("%_F005_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_ZMDPC_UPLOAD_COND_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_ZMDPC_UPLOAD_COND_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value",3,"","",DT_ZMDPC_UPLOAD_COND_3010_TABLECELL_SINGLE_VALUE_2,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_FAST_ENTRY",Convert4digitNumber(DT_GETCELLVALUE_OF_GRIDCELL_0_KBETR_EDIT_OUTPUt))
Call VerifyTableCellContent(2,"Amount","SAPMV13ATCTRL_FAST_ENTRY",Convert4digitNumber(DT_GETCELLVALUE_OF_GRIDCELL_1_KBETR_EDIT_OUTPUT))
'Call VerifyTableCellContent(3,"Amount","SAPMV13ATCTRL_FAST_ENTRY",Convert4digitNumber(DT_GETCELLVALUE_OF_GRIDCELL_2_KBETR_EDIT_OUTPUT)) '''--Commented as Competitor export excel updated with 2 sets of dat instead of 3 earlier



Call LogOff()
Call FinalStatus ()

