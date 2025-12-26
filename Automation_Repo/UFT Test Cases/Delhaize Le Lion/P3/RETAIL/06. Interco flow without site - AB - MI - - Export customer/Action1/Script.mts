'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_06. Interco flow without site - AB - MI --Export customer  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_06_ Interco flow - AB - MI"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

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

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''--------TransactionCode-	----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL02N_4004_OUTBOUND_DELIVERY,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TAXI_TABSTRIP_OVERVIEW", "Picking", False)
Call SetTableData("SAPMV50ATC_LIPS_PICK", "Picked Qty", 1, "", "", DT_VL02N_1104_TABLECELL_PICKED_QTY_0, False)
Call SetTableData("SAPMV50ATC_LIPS_PICK", "Picked Qty", 2, "", "", DT_VL02N_1104_TABLECELL_PICKED_QTY_1, False)
Call SetTableData("SAPMV50ATC_LIPS_PICK", "Picked Qty", 3, "", "", DT_VL02N_1104_TABLECELL_PICKED_QTY_2, False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
Call ClickButtonIfExist("Post Goods Issue   \(Shift\+F8\)",False)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
'Call VerifyStatusBar(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR)

'''''''--------TransactionCode-VF01----------''''
'
Call SetTcode(DT_VL02N_4004_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call VerifyTableCellContent(1,"Document", "SAPMV60ATCTRL_ERF_FAKT", DT_VL02N_0102_CHECK_TEXT_OF_TABLECELL_DOCUMENT_0)
Call SelectRowGuiTableByRow("SAPMV60ATCTRL_ERF_FAKT", 1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_OP_DOC_NO")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_VL02N_0102_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

'''''--------TransactionCode-VF03----------''''

Call SetTcode(DT_VL02N_0102_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Billing document", "VBRK-VBELN", "",DT_VL02N_0101_BILLING_DOCUMENT, False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Display document flow   \(Shift\+F7\)",False)
Call TakeScreenShot
Call ActivateItemGuiTree(0,"#1;#1;#1;#1","#1")
Call TakeScreenShot
'Call ClickButtonToolBar("&DETAIL",True)
Call GetGridContentByTitle("Accounting document.*","", "Doc.no.",1,"DT_OP_VL02N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM") ' Need to update
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo) 
''''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_VL02N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_VL02N_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_VL02N_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_VL02N_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonToolBar("&MB_FILTER", 0)
Call ClickButtonToolBar("&FIND", True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Branch account",True)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0, "ON", True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND", True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Trading Partner",True)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0, "ON", True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND", True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","G/L Account",True)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0, "ON", True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Column Set",0,22, True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot 
Call ClickButton("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "Posting Key",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "Account",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "Assignment",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 2, "Assignment",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)''Right(Date,4)+Mid(Date,4,2)+Left(Date,2))
Call VerifyGridCellContent("", 3, "Assignment",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)''Right(Date,4)+Mid(Date,4,2)+Left(Date,2))
Call VerifyGridCellContent("", 1, "Tax Code",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 1, "Business Area",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call VerifyGridCellContent("", 2, "Business Area",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)
Call VerifyGridCellContent("", 3, "Business Area",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_GSBER)
Call VerifyGridCellContent("", 1, "Profit Center",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "Profit Center",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 3, "Profit Center",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
'Call VerifyGridCellContent("", 1, "Cost Center",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOSTL)
'Call VerifyGridCellContent("", 2, "Cost Center",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
'Call VerifyGridCellContent("", 3, "Cost Center",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOSTL)
Call VerifyGridCellContent("", 1, "Alternative Account No.",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "Alternative Account No.",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 3, "Alternative Account No.",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
'Call VerifyGridCellContent("", 1, "Trading partner",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBUND)
'Call VerifyGridCellContent("", 2, "Trading partner",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_VBUND)
'Call VerifyGridCellContent("", 3, "Trading partner",0,DT_VL02N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VBUND)
'Call VerifyGridCellContent("", 1, "Trading partner",0,"")
'Call VerifyGridCellContent("", 2, "Trading partner",0,"")
'Call VerifyGridCellContent("", 3, "Trading partner",0,"")
Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type", "BKPF-BLART", "", lcase(DT_VL02N_1710_CHECK_TEXT_OF_DOCUMENT_TYPE), True)
Call TakeScreenShot
Call PressEnter()
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

