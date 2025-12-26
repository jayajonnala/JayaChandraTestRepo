'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_01_AUTOMATED IC LISTING BELGIUM 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "TC1_Test_PRE_01_AUTOMATED IC LISTING BELGIUM"
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-FB70----------''''


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB70_1000_COMPANY_CODE,True)
'Call TakeScreenShot
'Call PressEnter()
'Updated SetTextboxPopupIfExist as company code was defaulted 
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB70_1000_COMPANY_CODE)
Call PressEnter()
Call SetComboByKey("Transactn", DT_FB70_1200_TRANSACTN)
Call TakeScreenShot
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB70_0510_INVOICE_DATE),False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
Call SetComboByKey("Document Type", DT_FB70_0510_DOCUMENT_TYPE)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
Call SelectCheckbox("INVFO-XMWST",0,"ON", False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "", DT_FB70_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", 1, "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", 1, "", "", DT_FB70_0100_TABLECELL_PROFIT_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOC")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_CHECK_DOC)
Call ClickButtonIfExist("Cancel   \(F12\)",True)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
