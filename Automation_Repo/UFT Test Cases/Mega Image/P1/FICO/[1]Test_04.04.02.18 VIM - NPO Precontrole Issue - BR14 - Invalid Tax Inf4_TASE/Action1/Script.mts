
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.18 VIM - NPO Precontrole Issue - BR14 - Invalid Tax Inf_TASE
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


gstrTestCaseName = "Test_04.04.02.18 VIM"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode /OPT/VIM_7AX2----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)   
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)   
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Wait(2)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call SetTextbox("Document Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BLDAT","",DT_OPTVIM_7AX2_2002_DOCUMENT_DATE,False)
Call SetTextbox("Posting Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BUDAT","",DT_OPTVIM_7AX2_2002_POSTING_DATE,False)
Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call SetTextbox("Gross Inv Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-GROSS_AMOUNT","",DT_OPTVIM_7AX2_2002_GROSS_INV_AMOUNT,False)
Call SetTextbox("Expense Type","/OPT/VIM_BL_1RIDX_OCR_DATA-EXPENSE_TYPE","",DT_OPTVIM_7AX2_2002_EXPENSE_TYPE,False)
Call SetTextbox("Requester ID","/OPT/VIM_BL_1RIDX_OCR_DATA-REQUISITIONER","",DT_OPTVIM_7AX2_2002_REQUESTER_ID,False)
Call SetTextbox("VAT Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-VAT_AMOUNT","",DT_OPTVIM_7AX2_2002_VAT_AMOUNT,False)

Call TakeScreenShot()

Call ClickButton("SUBMIT   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call GetStatusBar("item2","DT_DOCUMENT_ID_OUTPUT")
Call GetStatusBar("item4","DT_WORKFLOW_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar("Document Number "&DT_DOCUMENT_ID_OUTPUT&" Workflow ID "&DT_WORKFLOW_ID_OUTPUT)


'------------------------------TCode /n/OPT/VIM_WP-------------------------------------------------

Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Wait(2)
Call SelectRadioButtonIfPopupExists("SPOPLI-SELFLAG","All Users View")
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Wait(2)


Call SetTextbox("Reference","H1_XBNR-LOW","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call TakeScreenShot()
Call ClickButton("Apply",False)
Call TakeScreenShot()


Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"Document Id",DT_DOCUMENT_ID_OUTPUT,False)
Call TakeScreenShot()

Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()

Call FindRowNumber("","Proc. Type",DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE,"DT_PROC_TYPE_ROW_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","224","Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT_0)
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","224","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS)


'Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Proc. Type",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE)
'Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT_0)
'Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS)

Call ClickButton("Exit   \(Shift\+F7\)",False)
Wait(2)


Call SelectTab("TAB_MAIN","Basic Data",False)
Wait(1)
Call TakeScreenShot()
Call SetTextbox("Cost Center","GH_IDX_APPLICATION->MS_IDX_HEADER-ZZKOSTL","",DT_OPTVIM_7AX2_9100_GH_IDX_APPLICATIONMS_IDX_HEADERZZKOSTL,False)
Call TakeScreenShot()



Call SelectTab("TAB_MAIN","Line Items",False)
Wait(1)
Call TakeScreenShot()
Call SetTableDataNoRef("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","G/L",1,DT_OPTVIM_7AX2_1210_TABLECELL_GL_0,False)
Call SetTableDataNoRef("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Amount",1,DT_OPTVIM_7AX2_1210_TABLECELL_AMOUNT_0,False)
Call SetTableDataNoRef("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Tax Code",1,DT_OPTVIM_7AX2_1210_TABLECELL_TAX_CODE_0,False)
'Call SetTableData("SAPLZTPVM_IDX_UITCTRL_ITEM_1210","Debit/Credit",1,"","","Debit",False)
Call SetTableDataNoRef("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Tax Rate",1,DT_OPTVIM_7AX2_1210_TABLECELL_TAX_RATE_0,False)
Call SetTableDataNoRef("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Company Code",1,DT_OPTVIM_7AX2_1210_TABLECELL_COMPANY_CODE_0,False)
Call SetTableDataNoRef("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Profit Center",1,DT_OPTVIM_7AX2_1210_TABLECELL_COST_CENTER_0,False)
Call PressEnter()



Call SelectTab("TAB_MAIN","Tax",False)
Wait(1)
Call SetTextbox("Tax Amount","GH_IDX_APPLICATION->MS_IDX_HEADER-VAT_AMOUNT","",DT_OPTVIM_7AX2_2002_VAT_AMOUNT,False)
Call SetTextbox("Tax Code","GH_IDX_APPLICATION->MS_IDX_HEADER-TAX_CODE","",DT_OPTVIM_7AX2_2002_OPTVIM_BL_1RIDX_OCR_DATATAX_CODE,False)
Call PressEnter()
Wait(5)
Call TakeScreenShot()



Call SelectTab("TAB_MAIN","Process",False)
Wait(1)
Call TakeScreenShot()


Call SelectTab("TAB_MAIN","Accounting",False)
Wait(1)
Call TakeScreenShot()
Call SetCombo("GH_IDX_APPLICATION->MS_IDX_HEADER-POST_DATE_DET","Manual Entry")
Call PressEnter()
Wait(2)
Call SetTextbox("Posting Date","GH_IDX_APPLICATION->MS_IDX_HEADER-BUDAT","",DT_ACCOUNTING_POSTING_DATE,False)
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()


Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)


Call VerifyStatusBar(DT_OPTVIM_7AX2_1000_CHECK_TEXT_OF_STATUSBAR)


Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()


Call FindRowNumber("","Proc. Type",DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE,"DT_PROC_TYPE_ROW_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

' VerifyGridCellContentByRefColumn(gridTitle, gridRowNumber, refColumnName, refCellValue, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContentByRefColumn("","","Process Type","224","Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT_OCC1)
Call VerifyGridCellContentByRefColumn("","","Process Type","224","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS_OCC1)


'Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Proc. Type",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE_OCC1)
'Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT_OCC1)
'Call VerifyGridCellContent("",DT_PROC_ROW_NUM,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS_OCC1)

Call ClickButton("Exit   \(Shift\+F7\)",False)
Wait(2)

'Click on Exit
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(2)
'Click on Exit
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

