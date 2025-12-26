

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCAOM01_002_DC_Inventory_Management_Destruction_fresh_DC
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

gstrTestCaseName = "Test_06DCAOM01_002_DC_Inventory_Management_Destruction_fresh_DC"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------ZMDIM_DESTR_STOCK-----------------------------


Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
'Call SetTextbox("Variant","V-LOW","",DT_ZMDIM_DESTR_STOCK_0841_SEARCH_TERM,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program .*","","Variant name",DT_ZMDIM_DESTR_STOCK_0841_SEARCH_TERM,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Choose   \(F2\)",True) 

 Call SetTextbox("Customer","P_KUNNR","",DT_ZMDIM_DESTR_STOCK_1000_CUSTOMER,False)
 Call SetTextbox("Reason for Movement","P_GRUND","",DT_ZMDIM_DESTR_STOCK_1000_REASON_FOR_MOVEMENT,False)
 Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_ZMDIM_DESTR_STOCK_1000_POSTING_DATE),False)
Call TakeScreenShot()

 Call SetTextbox("Site","S_WERKS-LOW","",DT_ZMDIM_DESTR_STOCK_1000_SITE,False)
 Call SetTextbox("Company Code","S_BUKRS-LOW","",DT_ZMDIM_DESTR_STOCK_1000_COMPANY_CODE,False)
 Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDIM_DESTR_STOCK_1000_ARTICLE,False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call PressEnter()     ' 

  Call  SetGridData("",1,"Pallet",FormatBlank(DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_0_PALLET),False)
Call SelectCellGuiGrid("",0,1,"Pallet",False)
Call SendKey("123456789")

Call PressEnter()     ' 

Call  SetGridData("",1,"Batch",FormatBlank(DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_0_BATCH),False)  
Call SelectCellGuiGrid("",0,1,"Batch",False)
Call SendKey("1")

Call PressEnter()     ' 

Call SelectRowGuiGridbyRowNo("",0,1,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Save Documents   \(F8\)",False) 
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call GetGridContentByTitle("",0,"MBLNR",1,"DT_CREATED_DOC_OUTPUT")
Call SelectCellGuiGrid("",0,1,"Created Document",False)
Call ClickCellGuiGrid("",0,"Created Document",1,"","",False)
Call TakeScreenShot()

'''''Added these lines of script to adapt the changes suggested in defect 20256
Call VerifyTextBoxContent("Article","GOITEM-MATNR",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ARTICLE_DESCRIPTION,False)
Call SelectTab("TS_GOITEM","Where",False)
Call TakeScreenShot()
Call VerifyTextBoxContent("Site","GOITEM-WERKS",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_SITE,False)
Call VerifyTextBoxContent("Storage Location","GOITEM-LGORT",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_SLOC,False)
Call VerifyTextBoxContent("Movement type","GOITEM-BWART",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_MVT,False)

'Call ClickButton("Accounting Documents   \(F7\)",False) 
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call TakeScreenShot()
Call ClickButton("OK_FI_DOC",False)
Call TakeScreenShot()

'Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'Call VerifyTextBoxContent("Material Description","MSEG-MATNR",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ARTICLE_DESCRIPTION,False)
'Call VerifyTextBoxContent("Site","MSEG-WERKS",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_SITE,False)
'Call VerifyTextBoxContent("SLoc","MSEG-LGORT",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_SLOC,False)
'Call VerifyTextBoxContent("MvT","MSEG-BWART",0,DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_MVT,False)

'Call ClickButton("Accounting Documents   \(F7\)",False) 
Call SelectRowGuiGridbyRowNo("Documents in Accounting",0,1,True)
Call DoubleClickGuiGridCell("Documents in Accounting","",1,"Object type text",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call VerifyGridCellContent("",1,"KTONR",0,DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"KTONR",0,DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"LOKKT",0,DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("",2,"LOKKT",0,DT_ZMDIM_DESTR_STOCK_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call LogOff()
Call FinalStatus ()


