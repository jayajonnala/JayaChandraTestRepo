	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_07.01.05.10.01 Goods Issue to internal order_Vehicle from BO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.01.05.10.01 Goods Issue to internal order_Vehicle from BO_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------

call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)


call ClickButtonIfExist("Open Detail Data",false)
Call SelectTab("TS_GOITEM","Where",False)
Call SetTextbox("Site","GOITEM-NAME1","",DT_MIGO_0325_GOITEMNAME1,False) 
Call SetTextbox("Storage Location","GOITEM-LGOBE","",DT_MIGO_0325_GOITEMLGOBE,False) 
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_0325_QTY_IN_UNIT_OF_ENTRY,False) 
Call PressEnter()  
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Quantity",False)
'''Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0325_QTY_IN_UNIT_OF_ENTRY,False) 
Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_QTY_IN_UNIT_OF_ENTRY,False) 
Call PressEnter()  
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Article",False)
Call SetTextbox("Article","GOITEM-MAKTX","",DT_MIGO_0310_QTY_IN_UNIT_OF_ENTRY,False)
Call PressEnter() 
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Where",False)
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_0325_QTY_IN_UNIT_OF_ENTRY,False) 
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Article",False)
Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False)
Call TakeScreenShot()
Call PressEnter() 
Call SelectTab("TS_GOITEM","Account Assignment",False)
Call SetTextbox("Order","COBL-AUFNR","",DT_MIGO_1002_ORDER,False) 
Call PressEnter() 
call ClickButton("Post Document   \(Shift\+F11\)",fALSE)
call ClickButtonIfExist("Continue   \(Enter\)",fALSE)
call GetStatusBar("item1","DT_Article_docnum_OUTPUT")
Call VerifyStatusBar("Article document " & DT_Article_docnum_OUTPUT &" posted")

Call LogOff()
Call FinalStatus ()


