'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Test_PRE3_MIRO_Invoice_TASE
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


gstrTestCaseName = "Test_PRE3_MIRO_Invoice_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DLL\FICO\TASE_DT_PRE3_MIRO_Invoice.xls"



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-ZFIAR_AFF_CLEARING----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)


Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("Invoice date","INVFO-BLDAT",0,ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR",0,DT_MIRO_0010_REFERENCE,False)
Call SelectCheckbox("INVFO-XMWST",0, "ON", False)
Call TakeScreenShot()
Call SetTextboxNoLabel("RM08M-EBELN",0,DT_MIRO_6211_RM08MEBELN,False)
Call PressEnter()

Call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",DT_MIRO_0010_AMOUNT,false)
Call GetTextboxValue("RM08M-DIFFERENZ", 0,"DT_NEG_AMOUNT_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_NEG_AMOUNT_OUTPUT",DT_NEG_AMOUNT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call PressEnter()
Call SetTableData("SAPLMR1MTC_MR1M","Tax Code","1","","",DT_MIRO_6310_TABLECELL_TAX_CODE_0,False)
Call PressEnter()


Call TakeScreenShot()

'''Call ClickButton("btn\[11\]",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("item1", "DT_DOCUMENT_NUMBER_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOCUMENT_NUMBER_OCC1_OUTPUT",DT_DOCUMENT_NUMBER_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Wait 5

'''Call VerifyStatusBar("Document no. "&DT_DOC_NUM_OUTPUT& " created")
Call VerifyStatusBar(DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()



'''SAPGuiSession("Session").SAPGuiWindow("Enter Incoming Invoice:").Maximize
'''SAPGuiSession("Session").SAPGuiWindow("Enter Incoming Invoice:").SAPGuiTable("Layout").SetCellData 1,"Tax Code","B1 (BE VAT DED ____% LOCAL PURCHASE TOWN RUBBISH BAGS)"

'*********************************************End Of Script*********************************************************************

