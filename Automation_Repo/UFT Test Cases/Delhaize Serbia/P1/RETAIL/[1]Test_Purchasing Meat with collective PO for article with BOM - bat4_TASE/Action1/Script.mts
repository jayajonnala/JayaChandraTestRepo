
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for article with BOM - bat4
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
gstrTestCaseName = "Test_Purchasing Meat with collective PO for article with BOM - bat4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\RETAIL\DT_Purchasing Meat with collective PO for article with BOM - bat4_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCRMENTE_DELIVERY_NOTE",(Cint(DT_INCRMENTE_DELIVERY_NOTE)+1))

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MIGO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call TakeScreenShot
Call SetComboByKey("GODYNPRO-REFDOC", DT_MIGO_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2000_GODYNPROPO_NUMBER, False)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call PressEnter()
Call TakeScreenShot

Call GetTableCellData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","","DT_QTY_IN_UNE_0_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_QTY_IN_UNE_0_OUTPUT",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_QTY_IN_UNE_0)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call SelectTab("TS_GOITEM","Batch",False)

Call SetGridData("",1,"Documentary Batch - Batch No.",DT_MIGO_0201_GRIDCELL_0_DB_NO_OCC2,False)
Call SetGridData("",1,"Qty in unit of entry",DT_MIGO_0201_GRIDCELL_0_QTY_IN_UN_OF_ENTRY_OCC4,False)
Call PressEnter() 
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call ClickButton("Back   \(F3\)",false)
Call GetStatusBar("Text","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)

Call LogOff()
Call FinalStatus ()



'*********************************************End Of Script*********************************************************************

