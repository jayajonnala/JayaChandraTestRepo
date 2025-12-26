

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.07.01 Manage Asset Trans Intra-Opco_Trans within CC
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_09.11.01.07.01 Manage Asset Trans Intra-Opco_Trans within CC
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.11.01.07.01 Manage Asset Trans Intra-Opco_Trans within CC"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode AS03----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 


Call VerifyTextBoxContent("Quantity","ANLA-MENGE",0,"0,000",False)

'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",2,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",3,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",5,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",6,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",7,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",8,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",9,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",10,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",11,"Change",0,"")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",12,"Change",0,"")
Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_CHECK_22_5))
Call VerifyGridCellContent("Transactions",2,"Asset Value Date",0,ConvertDate(DT_CHECK_22_6))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_CHECK_22_7)
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_CHECK_22_8)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_CHECK_22_9)
Call VerifyGridCellContent("Transactions",2,"Transaction Type",0,DT_CHECK_22_10)

Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",2,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",3,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",5,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",6,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",7,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",8,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",9,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",10,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",11,"Change",0,"")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",12,"Change",0,"")
Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_CHECK_20_5))
Call VerifyGridCellContent("Transactions",2,"Asset Value Date",0,ConvertDate(DT_CHECK_20_6))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_CHECK_20_7)
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_CHECK_20_8)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_CHECK_20_9)
Call VerifyGridCellContent("Transactions",2,"Transaction Type",0,DT_CHECK_20_10)


Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call SelectColumnGuiGrid("Transactions",0,"Transaction type name",False)
Call ClickButtonToolBar("&SORT_DSC",0)

Call SelectRowGuiGrid("Transactions",0,"Transaction type name","Retirmt transfer of curr-yr acquis.",False)
Call DoubleClickGuiGridCell("Transactions",0,2,"Transaction type name",False)
Call TakeScreenShot()
 

Call VerifyGridCellContent("",1,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call VerifyGridCellContent("",1,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("",2,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("",1,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("",2,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
Call VerifyGridCellContent("",1,"Text",0,Lcase(DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1))
Call VerifyGridCellContent("",2,"Text",0,Lcase(DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT_OCC1))


'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)


'----------------------Tcode AS03----------------------------

'Create Purchase Order
Call SetTcode(DT_AS03_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS03_0100_OKCD)


'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot()


'Click on Master Data
Call ClickButton("Master data   \(F7\)",False)


Call VerifyTextBoxContent("Quantity","ANLA-MENGE",0,DT_AS03_1140_CHECK_TEXT_OF_QUANTITY_OCC1,False)

'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_CHECK_22_15))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_CHECK_22_16)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_CHECK_22_17)

Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")


Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Change",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",1,"Year-end",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE_OCC3)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Change",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC1)
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation",4,"Year-end",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC3)
Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_CHECK_22_15))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_CHECK_22_16)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_CHECK_22_17)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")


Call SelectRowGuiGrid("Transactions",0,"Transaction type name","Acquiring transfer of curr-yr acquis.",False)
Call DoubleClickGuiGridCell("Transactions",0,1,"Transaction type name",False)
Call TakeScreenShot()
 

Call VerifyGridCellContent("",1,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call VerifyGridCellContent("",1,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("",2,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("",1,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("",2,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)
Call VerifyGridCellContent("",1,"Text",0,Lcase(DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1))
Call VerifyGridCellContent("",2,"Text",0,Lcase(DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT_OCC1))


'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

