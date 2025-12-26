'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Test_MR8M-Reverse PO Documents_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_MR8M-Reverse PO Documents"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DLL\FICO\TASE_DT_MR8M-Reverse PO Documents.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-MR8M----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Invoice Document No\.", "RBKPV-BELNR", 0, DT_MR8M_0300_INVOICE_DOCUMENT_NO, False)
Call SetTextbox("Fiscal Year","RBKPV-GJAHR",0,DT_MR8M_0300_FISCAL_YEAR,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD",0,DT_MR8M_0300_REVERSAL_REASON,False)
Call SetTextbox("Posting Date","G_BUDAT",0,ConvertDate(DT_MR8M_0300_POSTING_DATE),False)

Call TakeScreenShot

Call ClickButton("Display Document   \(F2\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call TakeScreenShot()

Call GetTableCellData("SAPLMR1MTC_MR1M", "Purchase order", 1, "Order Unit", "AU", "DT_MR8M_6310_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MR8M_6310_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0_OUTPUT",DT_MR8M_6310_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

'''Call ClickButton("btn\[3\]",False)
Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()

Call ClickButton("Reverse   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call GetStatusBar("item1", "DT_DOC_NO_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)

Wait 5
'''Call VerifyStatusBar("Document reversed with no."&DT_DOC_NO_OUTPUT&":Please Manually clear FI Documents")
Call VerifyStatusBar(DT_MR8M_0300_CHECK_TEXT_OF_STATUSBAR)


''--------TransactionCode-ME23N----------''''

Call SetTcode(DT_MR8M_0300_OKCD)
Call PressEnter()
Call TakeScreenShot()
'''Call ClickButton("btn\[17\]",False)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextBox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_MR8M_0003_PUR_ORDER,True)
'Call PressEnter()
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call TakeScreenShot()
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot()

Call SelectColumnGuiGrid("",0,DT_MR8M_0841_SEARCH_TERM,False)
Call Click204ButtonToolBar("&MB_FILTER",1)
Call TakeScreenShot()
Call ClickButton("Multiple selection",True)

Call SetTableData("SAPLALDBSINGLE","Single Value",1,"","",DT_MR8M_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single Value",2,"","",DT_MR8M_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(Enter\)",True)

Call VerifyGridCellContent("",3,"MENGE",0,DT_MR8M_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MENGE)
Call VerifyGridCellContent("",3,"DMBTR",0,DT_MR8M_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus()

'*********************************************End Of Script*********************************************************************
