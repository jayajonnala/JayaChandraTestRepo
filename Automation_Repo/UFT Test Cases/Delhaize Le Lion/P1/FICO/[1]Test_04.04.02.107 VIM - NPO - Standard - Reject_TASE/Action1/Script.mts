
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :AT_04.04.02.107 VIM - NPO - Standard - Reject
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



gstrTestCaseName = "AT_04.04.02.107 VIM - NPO - Standard - Reject"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'all LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Call TakeScreenShot()

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Document Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BLDAT","",DT_OPTVIM_7AX2_2002_DOCUMENT_DATE,False)
'Call SetTextbox("Vendor Number","/OPT/VIM_BL_1RIDX_OCR_DATA-LIFNR","",DT_OPTVIM_7AX2_2002_VENDOR_NUMBER,False)
Call SetComboByKey("/OPT/VIM_BL_1RIDX_OCR_DATA-POST_DATE_DET",DT_OPTVIM_7AX2_2002_POSTING_DATE_OCC1)

Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",FormatBlank(DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE),False)
Call SetTextbox("Gross Inv Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-GROSS_AMOUNT","",FormatBlank(DT_OPTVIM_7AX2_2002_GROSS_INV_AMOUNT),False)
Call SetTextbox("Expense Type","/OPT/VIM_BL_1RIDX_OCR_DATA-EXPENSE_TYPE","",FormatBlank(DT_OPTVIM_7AX2_2002_EXPENSE_TYPE),False)
Call SetTextbox("Requestor E-mail","/OPT/VIM_BL_1RIDX_OCR_DATA-EMAIL_ID","",FormatBlank(DT_OPTVIM_7AX2_2002_REQUESTOR_EMAIL),False)
Call SetTextbox("Requester ID","/OPT/VIM_BL_1RIDX_OCR_DATA-REQUISITIONER","",FormatBlank(DT_OPTVIM_7AX2_2002_REQUESTER_ID),False)

Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)

Call TakeScreenShot()
Call ClickButton("SUBMIT   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call GetStatusBar("item2","DT_OPTVIM_7AX2_0002_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
DT_DOCUMENT_ID = DT_OPTVIM_7AX2_0002_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT
Call GetStatusBar("item4","DT_OPTVIM_7AX2_0002_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call VerifyStatusBar(DT_OPTVIM_7AX2_0002_CHECK_TEXT_OF_STATUSBAR)

'--------------------------------------------  /n/opt/vim_wp----------------------------------------------

Call SetTcode(DT_OPTVIM_7AX2_0002_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Id","H1_DOID-LOW","",DT_DOCUMENT_ID,False)
Call ClickButton("BT_H_APPLY",False)
Call ClickButton("Switch Work View   \(Shift\+F6\)",False)
Call TakeScreenShot()
Call SelectRadioButton("SPOPLI-SELFLAG","All Users View",True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Hide Detail Pane   \(Ctrl\+F2\)",False)
Call ClickButton("Refresh   \(F5\)",False)
Wait 5

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False)
Call ClickButtonIfExist("Hide Detail Pane   \(Ctrl\+F2\)",False)
Call SelectTab("TAB_MAIN","Process",False)
Call SelectTab("TAB_MAIN","Basic Data",False)
''Call SetTextboxNoLabel("GH_IDX_APPLICATION->MS_IDX_HEADER-ZZKOSTL",0,FormatBlank(DT_OPTVIM_7AX2_9100_GH_IDX_APPLICATIONMS_IDX_HEADERZZKOSTL),False)
''If DT_OPTVIM_7AX2_9100_GH_IDX_APPLICATIONMS_IDX_HEADERZZKOSTL <> "" Then
Call SetTextboxNoLabel("GH_IDX_APPLICATION->MS_IDX_HEADER-ZZKOSTL",0,FormatBlank(DT_OPTVIM_7AX2_9100_GH_IDX_APPLICATIONMS_IDX_HEADERZZKOSTL),False)
Call SelectTab("TAB_MAIN","Line Items",False)

Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","G/L",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_GL_ACC_0,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Tax Code",1,"","",DT_TAX_CODE,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Amount",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_AMOUNT_0,False)
Call SetTableData("/OPT/SAPLVIM_IDX_UITCTRL_ITEM_1210","Cost Center",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_COST_CTR_0,False)
Call PressEnter()
''End If
''If DT_ACCOUNTING_HEADER_EMAIL <> "" Then
''	Call SetTextbox("Requester E-mail","GH_IDX_APPLICATION->MS_IDX_HEADER-EMAIL_ID","",FormatBlank(DT_ACCOUNTING_HEADER_EMAIL),False)
''End If
Call PressEnter()  
Call PressEnter()  
Call SelectTab("TAB_MAIN","Accounting",False)
Call PressEnter()  
Call PressEnter() 
Call ClickButton("btn\[11\]",False)

Call ClickButtonIfExist("Show Options   \(Ctrl\+F12\)",False)
'''Call ClickCellGuiGrid("",0,"Option Short Text",2,"","",False)
Call ClickCellGuiGrid("",0,"#1",2,"","",False) '' --- #1 is used to click on any first column as option short text column name is no more coming
 Wait 25
Call ClickButton("BT_H_APPLY",False)
wait 5
Call ClickButton("Refresh   \(F5\)",False)
'**************************** Updated below script - based on the Zephyr defect 30608- inputs for Potential rent invoice isue*************
Call TakeScreenShot()
Call GetGridContentByTitle("All Inbox.*",0,"Exception Reason",1,"DT_OPTVIM_7AX2_2004_TABLECELL_EX_REA_OUTPUT")
If DT_OPTVIM_7AX2_2004_TABLECELL_EX_REA_OUTPUT =  DT_OPTVIM_7AX2_2004_TABLECELL_EX_REA Then
	Call TakeScreenShot()
	Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
	Call ClickButton("Simulate and Bypass Business Rules   \(Ctrl\+Shift\+F9\)",False)
	Call TakeScreenShot()
	'ClickCellGuiGrid(gridTitle,gridIndex,columnName,rowNumber,columnNameRef,gridValRef,blnIsItPopup)
	Call ClickCellGuiGrid( "","","Activate/Bypass","9","Business Rule","Potential Rent Invoice (NPO)",False)
	Call TakeScreenShot()
	Call ClickButtonIfExist("Yes",True)
	Call SetTextArea(DT_OPTVIM_7AX2_2004_BYPASS_COMMENTS)
	Call TakeScreenShot()
	Call ClickButton("Save   \(F2\)",True)
	Call ClickButton("Exit   \(Shift\+F7\)",True)
	Call TakeScreenShot()
	Call ClickCellGuiGrid("",0,"Option",2,"","",False)
	 Wait 25
	Call ClickButton("BT_H_APPLY",False)
	wait 5
	Call ClickButton("Refresh   \(F5\)",False)
	Call TakeScreenShot()
	Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
	Call TakeScreenShot()
End If
'*********************************************************
Call VerifyWindowValue(DT_OPT_VIM_CHECK_1ST_APPROVER)

'Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","G/L Acc",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_GL_ACC_0,False)
'Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","Cost Ctr",1,"","",DT_OPTVIM_7AX2_2004_TABLECELL_COST_CTR_0,False)
'Call SetTableData("/ORS/SAPL000007_APPR_SCREENTC_LINE","Tax Code",1,"","",DT_TAX_CODE,False)
'Call PressEnter() 

Call ClickButton("Approve",False)
Call SetTextArea(DT_OPTVIM_7AX2_2010_TEXTEDIT_SHELL)
Call ClickButton("APPR_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_2ND_APPROVER)


Call ClickButton("Approve",False)
Call SetTextArea(DT_OPTVIM_7AX2_2010_TEXTEDIT_SHELL)
Call ClickButton("APPR_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)

Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_3RD_APPROVER)


Call ClickButton("Reject",False)
Call SetTextArea(DT_OPTVIM_7AX2_2020_TEXTEDIT_SHELL_OCC1)
Call TakeScreenShot()
Call ClickButton("REJECT_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_2ND_REJECTOR)

Call ClickButton("Reject",False)
Call SetTextArea(DT_OPTVIM_7AX2_2020_TEXTEDIT_SHELL_OCC2)
Call TakeScreenShot()
Call ClickButton("REJECT_BUTTON",False)
 Wait 5
Call ClickButton("Refresh   \(F5\)",False)
Call ClickCellGuiGrid("All Inbox.*",0,"Execute",1,"","",False) 
Call VerifyWindowValue(DT_OPT_VIM_CHECK_3RD_REJECTOR)

Call ClickButton("Reject",False)
Call SetTextArea(DT_OPTVIM_7AX2_2020_TEXTEDIT_SHELL_OCC3)
Call TakeScreenShot()
Call ClickButton("REJECT_BUTTON",False)
 Wait 5

Call ClickButton("Back   \(F3\)",False)

'''--------------------------------------------   /n/opt/vim_va2----------------------------------------------
Call SetTcode(DT_OPTVIM_7AX2_1000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Document Processing Number","S_DOCID-LOW","",DT_OPTVIM_7AX2_1000_DOCUMENT_PROCESSING_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call ClickButton("Refresh   \(F5\)",False)
Call VerifyGridCellContent("Results \(1 Hit\)",1,"OVERALL_STATUS_TEXT",0,DT_OPTVIM_7AX2_2000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_OVERALL_STATUS_TEXT)

Call LogOff()
Call FinalStatus ()


