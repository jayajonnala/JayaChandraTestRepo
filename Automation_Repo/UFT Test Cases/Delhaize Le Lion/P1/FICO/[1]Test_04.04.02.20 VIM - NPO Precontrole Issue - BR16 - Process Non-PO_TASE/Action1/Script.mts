
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.20 VIM - NPO Precontrole Issue - BR16 - Process Non-PO
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


gstrTestCaseName = "Test_04.04.02.20 VIM BR16NonPO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_NEW",Cint(DT_INCREMENT_NEW)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Call TakeScreenShot()

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call VerifyTextBoxContent("Arch\.","S_ARCHIV_ID",0,DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Document Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BLDAT","",DT_OPTVIM_7AX2_2002_DOCUMENT_DATE,False)
Call SetTextbox("Posting Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BUDAT","",DT_OPTVIM_7AX2_2002_POSTING_DATE,False)

Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",FormatBlank(DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE),False)
Call SetTextbox("Gross Inv Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-GROSS_AMOUNT","",FormatBlank(DT_OPTVIM_7AX2_2002_GROSS_INV_AMOUNT),False)
Call SetTextbox("Expense Type","/OPT/VIM_BL_1RIDX_OCR_DATA-EXPENSE_TYPE","",FormatBlank(DT_OPTVIM_7AX2_2002_EXPENSE_TYPE),False)
Call SetTextbox("Requestor E-mail","/OPT/VIM_BL_1RIDX_OCR_DATA-EMAIL_ID","",FormatBlank(DT_OPTVIM_7AX2_2002_REQUESTOR_EMAIL),False)
Call SetTextbox("Requester ID","/OPT/VIM_BL_1RIDX_OCR_DATA-REQUISITIONER","",FormatBlank(DT_OPTVIM_7AX2_2002_REQUESTER_ID),False)
Call SetTextbox("VAT Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-VAT_AMOUNT","",FormatBlank(DT_OPTVIM_7AX2_2002_VAT_AMOUNT),False)
Call SetTextboxNoLabel("/OPT/VIM_BL_1RIDX_OCR_DATA-TAX_CODE",0,FormatBlank(DT_OPTVIM_7AX2_2002_OPTVIM_BL_1RIDX_OCR_DATATAX_CODE),False)
Call SetTextboxNoLabel("/OPT/VIM_BL_1RIDX_OCR_DATA-TAX_RATE",0,FormatBlank(DT_OPTVIM_7AX2_2002_OPTVIM_BL_1RIDX_OCR_DATATAX_RATE),False)
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call SelectCheckbox("/OPT/VIM_BL_1RIDX_OCR_DATA-CREDIT_MEMO",0,DT_OPTVIM_7AX2_2002_CREDIT_MEMO,False)

Call TakeScreenShot()
Call ClickButton("SUBMIT   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call GetTextStatusBar("DT_OPTVIM_7AX2_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT")
Call VerifyStatusBar(DT_OPTVIM_7AX2_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------

'
Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Reference","H1_XBNR-LOW","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call ClickButton("BT_H_APPLY",False)
Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call TakeScreenShot()
Call SelectRadioButton("SPOPLI-SELFLAG","All Users View",True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Hide Detail Pane   \(Ctrl\+F2\)",False)
Call ClickButtonToolBar("&MB_FILTER",0)
Call SelectCellGuiGrid("Column Set",0,5,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Document Id","%%DYN001-LOW","",DT_OPTVIM_7AX2_1105_DOCUMENT_ID,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)",True)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call TakeScreenShot()
 Wait 5

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call SelectTab("TAB_MAIN","Process",False)
Call SelectTab("TAB_MAIN","Basic Data",False)
Call SetTextboxNoLabel("GH_IDX_APPLICATION->MS_IDX_HEADER-ZZKOSTL",0,FormatBlank(DT_OPTVIM_7AX2_9100_GH_IDX_APPLICATIONMS_IDX_HEADERZZKOSTL),False)
Call PressEnter()  
Call PressEnter()  
Call SelectTab("TAB_MAIN","Accounting",False)
Call PressEnter()  
Call SelectTab("TAB_MAIN","Line Items",False)
Call TakeScreenShot()
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","G/L",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_GL_ACC_0_NEW,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Amount",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_AMOUNT_0,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Cost Center",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_COST_CTR_0,False)
Call PressEnter() 
Call TakeScreenShot()
Call ClickButton("btn\[11\]",False)

Call ClickButtonIfExist("Show Options   \(Ctrl\+F12\)",False)
'''''Call ClickCellGuiGrid("",0,"Option Short Text",2,"","",False)
Call ClickCellGuiGrid("",0,"Option",2,"","",False)
Call ClickButton("Refresh   \(F5\)",False)
 Wait 25
''Call ClickButton("BT_H_APPLY",False)
''wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_1ST_APPROVER)

Wait 2
''Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","Cost Ctr",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_COST_CTR_0,False)
''Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","G/L Acc",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_GL_ACC_0_NEW,False)

Call PressEnter() 
Call ClickButton("Approve",False)
Call ClickButton("APPR_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_2ND_APPROVER)


Call ClickButton("Approve",False)
Call ClickButton("APPR_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_3RD_APPROVER)


Call ClickButton("Approve",False)
Call ClickButton("APPR_BUTTON",False)
Call ClickButton("Back   \(F3\)",False)
'
''--------------------------------------------   /n/opt/vim_va2----------------------------------------------
Call SetTcode(DT_OPTVIM_7AX2_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Processing Number","S_DOCID-LOW","",DT_OPTVIM_7AX2_1000_DOCUMENT_PROCESSING_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Refresh   \(F5\)",False)
Call VerifyGridCellContent("Results \(1 Hit\)",1,"DOCID",0,DT_OPTVIM_7AX2_2000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCID)
Call VerifyGridCellContent("Results \(1 Hit\)",1,"OVERALL_STATUS_TEXT",0,DT_OPTVIM_7AX2_2000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_OVERALL_STATUS_TEXT)
Call ClickCellGuiGrid("Results \(1 Hit\)",0,"Document Id",1,"","",False)
Call VerifyComboBoxValue("Trans.*" ,UCASE(DT_OPTVIM_7AX2_9100_CHECK_VALUE_OF_TRANSEVENT))
Call ClickButton("Cancel   \(F12\)",True)

Call LogOff()
Call FinalStatus ()




