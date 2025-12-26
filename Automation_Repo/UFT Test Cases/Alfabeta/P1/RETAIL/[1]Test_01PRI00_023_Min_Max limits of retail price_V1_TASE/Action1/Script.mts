

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_023_Min_Max limits of retail price_V1
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

gstrTestCaseName = "Test_01PRI00_023_Min_Max limits of retail price_V1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_023_Upload_Vendor_prices_via_xls_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  ZMDPC_UPLOAD_COND------------------------------------------------

Call RefreshExcelSheet(DT_ZMDPC_UPLOAD_COND_1000_FILE_NAME_OCC2) 'Added this line of code to refresh the Uplaod excel sheet to avoid Excel loading isues.

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program ZMDPC_UPLOAD_PRICES_FROM_FILE.*","","Variant name",DT_VARIANT,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True)

Call SetTextbox("File name","P_FILE","",DT_ZMDPC_UPLOAD_COND_1000_FILE_NAME_OCC2,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call VerifyGridCellContent("",1,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT)
Call VerifyGridCellContent("",2,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LIGHT)
Call VerifyGridCellContent("",3,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LIGHT)
Call VerifyGridCellContent("",3,"Message Text",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MSGTX)
Call GetGridContent("",0,"KBETR_EDIT",1,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_0_KBETR_OUTPUT")
'''Call SetGridData("",1,"KBETR_EDIT",DT_ZMDPC_UPLOAD_COND_500_GRIDCELL_0_PRICEEDIT,False)
Call GetGridContent("",0,"KBETR_EDIT",2,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_1_KBETR_OUTPUT")
Call VerifyGridCellContent("",1,"Valid From",0,ConvertDate(DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAB))
Call VerifyGridCellContent("",1,"Valid To",0,ConvertDate(DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATBI))

Call  ClickButton("Select All   \(F5\)",False)
Call  ClickButton("Create Conditions   \(F8\)",False)
Call VerifyGridCellContent("",1,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIGHT_OCC2)
Call VerifyGridCellContent("",2,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LIGHT_OCC2)
Call VerifyGridCellContent("",3,"Exception",0,DT_ZMDPC_UPLOAD_COND_500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LIGHT_OCC2)
Call GetGridContent("",0,"MATNR",1,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")
Call GetGridContent("",0,"LIFNR",1,"<NA>","<NA>","DT_ZMDPC_UPLOAD_COND_500_GETCELLVALUE_OF_GRIDCELL_0_LIFNR_OUTPUT")

'--------------------------------------------  VKP5------------------------------------------------
Call SetTcode(DT_VKP5_TRANSACTION) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VKP5_TRANSACTION)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_VARIANT_VKP5,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 

Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_ARTICLE_1_CHECK,True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_ARTICLE_2_CHECK,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SetGridData("",1," ENDPR",DT_NEW_FINAL_PRICE,False)
Call SetGridData("",2," ENDPR",DT_NEW_FINAL_PRICE,False)
Call VerifyGridCellContent("",1,"Article",0,DT_ARTICLE_1_CHECK)
Call VerifyGridCellContent("",2,"Article",0,DT_ARTICLE_2_CHECK)
Call VerifyGridCellContent("",2,"ENDPR",0,DT_NEW_FINAL_PRICE)
Call SelectAllRowGuiGrid("",0,False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call TakeScreenShot()

Call ClickButtonIfExist("Enter   \(F5\)",True)
Call ClickButtonIfExist("Enter   \(F5\)",True)

Call VerifyGridCellContentIfExist("",1,"F003",0,DT_MSG_TXT1)
Call VerifyGridCellContentIfExist("",2,"F003",0,DT_MSG_TXT2)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

'--------------------------------------------  VKP5------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ZMDPC_UPLOAD_COND_500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Purchasing Organization","I_EKORG","",DT_ZMDPC_UPLOAD_COND_1000_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Vendor","I_LIFNR","",DT_ZMDPC_UPLOAD_COND_1000_VENDOR,False)
Call SetTextbox("Article","I_MATNR","",DT_ZMDPC_UPLOAD_COND_1000_ARTICLE,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyWindowValue("Conditions: General Overview")

Call LogOff()
Call FinalStatus ()



