'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'.................Test_E10G06P01S02V12 Post GL Account document   
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E10G06P01S02V12 Post GL Acc"
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

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''--------TransactionCode-FB50L----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


Call PressEnter()
Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50L_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","","LU54",True)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50L_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FB50L_1010_DOCUMENT_TYPE,False)
Call SetTextbox("Posting Date","ACGL_HEAD-BUDAT","",ConvertDate(DT_FB50L_1010_POSTING_DATE),False)
Call SetTextbox("Ledger Grp"," ACGL_HEAD-LDGRP","",DT_FB50L_1010_LEDGER_GRP,False)
Call SetTextbox("Doc.Header Text"," ACGL_HEAD-BKTXT","",DT_FB50L_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Currency","ACGL_HEAD-WAERS","",DT_FB50L_1010_CURRENCY,False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50L_1010_REFERENCE,False)
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",ConvertDate(DT_FB50L_1010_DOCUMENT_DATE),False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB50L_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 2, "", "",DT_FB50L_0100_TABLECELL_GL_ACCT_1, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 3, "", "",DT_FB50L_0100_TABLECELL_GL_ACCT_2, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 4, "", "",DT_FB50L_0100_TABLECELL_GL_ACCT_3, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 1, "", "",DT_FB50L_0100_TABLECELL_DC_0, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 2, "", "",DT_FB50L_0100_TABLECELL_DC_1, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 3, "", "",DT_FB50L_0100_TABLECELL_DC_2, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", 4, "", "",DT_FB50L_0100_TABLECELL_DC_3, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 2, "", "",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 3, "", "",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_2, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 4, "", "",DT_FB50L_0100_TABLECELL_AMOUNT_IN_DOCCURR_3, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 1, "", "",DT_FB50L_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 2, "", "",DT_FB50L_0100_TABLECELL_TEXT_1, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 3, "", "",DT_FB50L_0100_TABLECELL_TEXT_2, False)
Call SetTableData("SAPLFSKBTABLE", "Text", 4, "", "",DT_FB50L_0100_TABLECELL_TEXT_3, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "",DT_FB50L_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", 2, "", "",DT_FB50L_0100_TABLECELL_PROFIT_CENTER_1, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", 3, "", "",DT_FB50L_0100_TABLECELL_PROFIT_CENTER_2, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 4, "", "",DT_FB50L_0100_TABLECELL_COST_CENTER_3, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Wait 5
Call VerifyStatusBar(DT_FB50L_1001_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_FB50L_1010_DOCUMENT_TYPE_OCC1,False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOCUMENT_NO")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_FB50L_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call ClickButtonIfExist("Cancel   \(F12\)",False)
Wait 5
Call ClickButtonIfExist("Yes",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call TakeScreenShot
'Call ClickButtonIfExist("Yes",False)
'''''''--------TransactionCode-/FAGLL03----------''''

Call SetTcode(DT_FB50L_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
         
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FB50L_1000_ACCOUNT_NUMBER,False)
Call ClickButton("Choose Ledger   \(Ctrl\+F3\)",False)
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FB50L_0300_LEDGER_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree("","General Ledger Line Items;Document Number")
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_DOC_NUM,False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Document Number", 0, DT_DOC_NUM)
'''''Call VerifyifGuiLabelExists_ByIndex(DT_DOC_NUM, 1)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Choose Ledger   \(Ctrl\+F3\)",False)
Call SetTextbox("Ledger","SVALD-VALUE","",DT_FB50L_0300_LEDGER,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_FAGLL03_0120_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
