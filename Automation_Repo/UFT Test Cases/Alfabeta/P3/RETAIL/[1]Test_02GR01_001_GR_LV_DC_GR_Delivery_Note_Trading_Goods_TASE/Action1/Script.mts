

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02GR01_001_GR_LV_DC_GR_Delivery_Note_Trading_Goods
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

gstrTestCaseName = "Test_02GR01_001_GR_LV_DC_GR_Delivery_Note_Trading_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MIGO-----------------------------


'
call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
Call TakeScreenShot()
call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call TakeScreenShot()
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0110_POSTING_DATE),False)
Call TakeScreenShot()
Call ClickButtonIfExist("MIGO_OK_GO",False)

Call TakeScreenShot()
Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call PressEnter() 

Call SelectTab("TS_GOITEM","Quantity",False)
'''Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_ORDER_QUANTITY1,False)
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_ORDER_QUANTITY,False)
Call SelectTab("TS_GOITEM","Batch",False)
'SAPGuiSession("Session_2").SAPGuiWindow("Goods Receipt Purchase_2").Maximize
Call TakeScreenShot()
Call SetGridData("",1,"Documentary Batch - Batch No.",DT_MIGO_0201_GRIDCELL_0_DB_NO,False)
Call SetGridData("",1,"Qty in unit of entry",DT_MIGO_0201_GRIDCELL_0_QTY_IN_UN_OF_ENTRY,False)
Call ClickCellGuiGrid("",0,"Classification",1,"","",False)
Call PressEnter() 

Call ClickButtonIfExist("Check Entries   \(F7\)",False)
Call SetTableData("SAPLCTMSCHARS_S","Value","1","","",DT_BATCH_NUM_ART,False) 
Call SetTableData("SAPLCTMSCHARS_S","Value","2","","",DT_AB_ORIGIN_NUM,False) 
Call SetTableData("SAPLCTMSCHARS_S","Value","3","","",DT_AB_VENDOR_BATCH,False) 
Call ClickButton("Back   \(F3\)",False)  
'Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call PressEnter() 

Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call ClickButton("Back   \(F3\)",False)  
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_ARTICLE_DOC_OUTPUT")
Call VerifyStatusBar("Article document "& DT_ARTICLE_DOC_OUTPUT &" posted")

Call LogOff()
Call FinalStatus ()





