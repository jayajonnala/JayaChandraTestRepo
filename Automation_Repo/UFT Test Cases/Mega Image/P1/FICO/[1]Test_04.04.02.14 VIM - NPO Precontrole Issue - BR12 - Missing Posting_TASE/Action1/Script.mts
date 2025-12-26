
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.14 VIM - NPO Precontrole Issue - BR12 - Missing Posting
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

gstrTestCaseName = "Test_04.04.02.14 BR12"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode /OPT/VIM_7AX2----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)   
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)   
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Wait(2)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN",""," ",False)
Call SetTextbox("Posting Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BUDAT",""," ",False)

Call SetCombo("/OPT/VIM_BL_1RIDX_OCR_DATA-POST_DATE_DET","Manual Entry")
Call TakeScreenShot()

Call GetTextboxValue("/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR",0,"DT_VENDOR_REFERENCE_NO",False)

Call ClickButton("SUBMIT   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call GetStatusBar("item2","DT_DOCUMENT_ID_OUTPUT")
Call GetStatusBar("item4","DT_WORKFLOW_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar("Document Number "&DT_DOCUMENT_ID_OUTPUT&" Workflow ID "&DT_WORKFLOW_ID_OUTPUT)


''''------------------------------TCode /n/OPT/VIM_WP-------------------------------------------------

Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
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

Call SelectTab("TAB_MAIN","Accounting",False)
Wait(1)
Call SetTextbox("Posting Date","GH_IDX_APPLICATION->MS_IDX_HEADER-BUDAT",""," ",False)
Call PressEnter()
Wait(2)
Call TakeScreenShot()


Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()

Call FindRowNumber("","Proc. Type",DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE,"DT_PROC_TYPE_ROW_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","913","Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT)
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","913","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS)


'Call VerifyGridCellContent("",DT_ROWNUM,"Proc. Type",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE)
'Call VerifyGridCellContent("",DT_ROWNUM,"Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT)
'Call VerifyGridCellContent("",DT_ROWNUM,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS)

Call ClickButton("Exit   \(Shift\+F7\)",False)
Wait(2)

Call SelectTab("TAB_MAIN","Accounting",False)
Wait(1)
Call SetTextbox("Posting Date","GH_IDX_APPLICATION->MS_IDX_HEADER-BUDAT","",DT_OPTVIM_7AX2_1300_POSTING_DATE,False)
Call PressEnter()
Wait(5)
Call TakeScreenShot()


Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
Wait(1)
Call TakeScreenShot()


Call FindRowNumber("","Proc. Type",DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE_OCC1,"DT_PROC_TYPE_ROW_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","913","Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT_OCC1)
Call VerifyGridCellContentByRefColumn("","","PROCESS_TYPE","913","Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS_OCC1)


'Call VerifyGridCellContent("",DT_ROWNUM,"Proc. Type",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PROCESS_TYPE_OCC1)
'Call VerifyGridCellContent("",DT_ROWNUM,"Business Rule",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_OBJTXT_OCC1)
'Call VerifyGridCellContent("",DT_ROWNUM,"Status",0,DT_OPTVIM_7AX2_2004_CHECK_GETCELLVALUE_OF_GRIDCELL_10_STATUS_OCC1)

Call ClickButton("Exit   \(Shift\+F7\)",False)
Wait(2)


Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)


Call VerifyStatusBar(DT_OPTVIM_7AX2_1000_CHECK_TEXT_OF_STATUSBAR)


Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(2)
Call ClickButton("Exit   \(Shift\+F3\)",False) 
Wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
