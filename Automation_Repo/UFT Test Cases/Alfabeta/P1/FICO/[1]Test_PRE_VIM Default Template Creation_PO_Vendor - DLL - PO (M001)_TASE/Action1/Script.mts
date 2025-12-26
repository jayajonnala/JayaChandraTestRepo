

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_VIM Default Template Creation_PO_Vendor
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




gstrTestCaseName = "Test_PRE_VIM Default Template Creation_PO_Vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_008_Creation_of_SPC_for_competition_driven_articles_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

SAPGuiUtil.CloseConnections
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
'
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  ZTPVM_DOC_DET------------------------------------------------

Call ClickButton("Position\.\.\.",False)
Call SetTextbox("Company Code","SVALD-VALUE",0,DT_ZTPVM_DOC_DET_0300_COMPANY_CODE,True)
Call SetTextbox("Company Code","SVALD-VALUE",1,DT_ZTPVM_DOC_DET_0300_COMPANY_CODE_OCC1,True)
'Call SetTextbox("Vendor","SVALD-VALUE",0,DT_ZTPVM_DOC_DET_0300_VENDOR,True)
'Call SetTextbox("Vendor","SVALD-VALUE",1,DT_ZTPVM_DOC_DET_0300_VENDOR_OCC1,True)
Call SetTextbox("Supplier","SVALD-VALUE",0,DT_ZTPVM_DOC_DET_0300_VENDOR,True)
Call SetTextbox("Supplier","SVALD-VALUE",1,DT_ZTPVM_DOC_DET_0300_VENDOR_OCC1,True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()



'--------------------------------------------  xk03----------------------------------------------
Call SetTcode(DT_ZTPVM_DOC_DET_0001_OKCD) 
Call PressEnter()  
Call TakeScreenShot()' 
Call CheckTCodeScreen(DT_ZTPVM_DOC_DET_0001_OKCD)

Call SelectCheckbox("RF02K-D0110",0,DT_ZTPVM_DOC_DET_0101_ADDRESS,False)
Call SelectCheckbox("RF02K-D0120",0,DT_ZTPVM_DOC_DET_0101_CONTROL,False)
Call SelectCheckbox("RF02K-D0130",0,DT_ZTPVM_DOC_DET_0101_PAYMENT_TRANSACTIONS,False)
Call SelectCheckbox("WRF02K-D0380",0,DT_ZTPVM_DOC_DET_0101_CONTACT_PERSONS,False)
Call SelectCheckbox("RF02K-D0210",0,DT_ZTPVM_DOC_DET_0101_ACCOUNTING_INFO,False)
Call SelectCheckbox("RF02K-D0215",0,DT_ZTPVM_DOC_DET_0101_PAYMENT_TRANSACTIONS_OCC1,False)
Call SelectCheckbox("RF02K-D0220",0,DT_ZTPVM_DOC_DET_0101_CORRESPONDENCE,False)
Call SelectCheckbox("RF02K-D0610",0,DT_ZTPVM_DOC_DET_0101_WITHHOLDING_TAX,False)

Call SetTextbox("Supplier","RF02K-LIFNR","",DT_ZTPVM_DOC_DET_0101_VENDOR,False)
Call SetTextbox("Company Code","RF02K-BUKRS",0,DT_ZTPVM_DOC_DET_0101_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter()     ' 
Call TakeScreenShot()

Call VerifyTextBoxContent("Supplier","RF02K-LIFNR",0,DT_ZTPVM_DOC_DET_0111_CHECK_TEXT_OF_VENDOR,False)
Call GetspecialTextboxValue("ADDR1_DATA-NAME1",0,"DT_ZTPVM_DOC_DET_0301_CHECK_TEXT_OF_NAME_OUTPUT",False)
Call ClickButton("Next screen   \(F8\)",False)
Call GetTextboxValue("LFA1-STCEG",0,"DT_ZTPVM_DOC_DET_0120_CHECK_TEXT_OF_VAT_REG_NO_OUTPUT",False)
Call GetTextboxValue("LFA1-STCD1",0,"DT_GET_TAX_NUMBER1_OUTPUT",False)
Call GetTextboxValue("LFA1-STCD2",0,"DT_GET_TAX_NUMBER2_OUTPUT",False)
Call GetTextboxValue("LFA1-STCD3",0,"DT_GET_TAX_NUMBER3_OUTPUT",False)
Call TakeScreenShot()
Call ClickButton("Next screen   \(F8\)",False)
Call TakeScreenShot()

Call GetTableCellData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Ctry",1,"","","DT_ZTPVM_DOC_DET_0130_CHECK_TEXT_OF_TABLECELL_CTRY_0_OUTPUT",False)
Call GetTableCellData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Key",1,"","","DT_ZTPVM_DOC_DET_0130_CHECK_TEXT_OF_TABLECELL_BANK_KEY_0_OUTPUT",False)
Call GetTableCellData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Bank Account",1,"","","DT_ZTPVM_DOC_DET_0130_CHECK_TEXT_OF_TABLECELL_BANK_ACCOUNT_0_OUTPUT",False)
Call GetTableCellData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","IBANValue",1,"","","DT_ZTPVM_DOC_DET_0130_CHECK_TEXT_OF_TABLECELL_IBANVALUE_0_OUTPUT",False)
Call GetTableCellData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Name of Financial Institution",1,"","","DT_ZTPVM_DOC_DET_0130_CHECK_TEXT_OF_TABLECELL_NAME_OF_BANK_0_OUTPUT",False)
Call ClickCellTableByRowNo("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","IBAN",1,False)
Call TakeScreenShot()
Call GetTextboxValue("BNKA-SWIFT",0,"DT_ZTPVM_DOC_DET_0100_CHECK_TEXT_OF_SWIFTBIC_OUTPUT",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Next screen   \(F8\)",False)
Call TakeScreenShot()
Call GetTextboxValue("LFB1-ZTERM",0,"DT_ZTPVM_DOC_DET_0215_CHECK_TEXT_OF_PAYT_TERMS_OUTPUT",False)
Call VerifyTextBoxContent("Assign\.Grp","LFB1-ASSIGN_TEST",0,DT_ZTPVM_DOC_DET_0215_CHECK_TEXT_OF_ASSIGNGRP,False)


'--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE2) 
Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call ClickButton("Copy From",False)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Document Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BLDAT","",DT_OPTVIM_7AX2_2002_DOCUMENT_DATE,False)
Call SetTextbox("Posting Date","/OPT/VIM_BL_1RIDX_OCR_DATA-BUDAT","",DT_OPTVIM_7AX2_2002_POSTING_DATE,False)
Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE,False)
Call SetTextbox("Vendor Number","/OPT/VIM_BL_1RIDX_OCR_DATA-LIFNR","",DT_OPTVIM_7AX2_2002_VENDOR_NUMBER,False)
Call SetSpecialTextbox("Vendor Name","/OPT/VIM_BL_1RIDX_OCR_DATA-VEND_NAME","",DT_OPTVIM_7AX2_2002_VENDOR_NAME,False)
Call SetTextbox("Company Code","/OPT/VIM_BL_1RIDX_OCR_DATA-BUKRS","",DT_OPTVIM_7AX2_2002_COMPANY_CODE,False)
Call SetTextbox("Document Currency","/OPT/VIM_BL_1RIDX_OCR_DATA-WAERS","",DT_OPTVIM_7AX2_2002_DOCUMENT_CURRENCY,False)
Call SetTextbox("Gross Inv Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-GROSS_AMOUNT","",DT_OPTVIM_7AX2_2002_GROSS_INV_AMOUNT,False)
Call SetTextbox("Freight Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-FREIGHT_AMOUNT","",FormatBlank(DT_OPTVIM_7AX2_2002_FREIGHT_AMOUNT),False)
Call SetTextbox("Handling Charges","/OPT/VIM_BL_1RIDX_OCR_DATA-HANDLING_CHARGES","","",False)
Call SetTextbox("Expense Type","/OPT/VIM_BL_1RIDX_OCR_DATA-EXPENSE_TYPE","",DT_OPTVIM_7AX2_2002_EXPENSE_TYPE,False)
Call SetTextbox("Plant","/OPT/VIM_BL_1RIDX_OCR_DATA-WERKS","","",False)
Call SetTextbox("Requestor E-mail","/OPT/VIM_BL_1RIDX_OCR_DATA-EMAIL_ID","","",False)
Call SetTextbox("Requester ID","/OPT/VIM_BL_1RIDX_OCR_DATA-REQUISITIONER","","",False)
Call SetTextbox("Recepient Name","/OPT/VIM_BL_1RIDX_OCR_DATA-RECIPIENT_NAME","","",False)
Call SetTextbox("PO List","/OPT/VIM_BL_1RIDX_OCR_DATA-PO_LIST","","",False)

Call SetTextbox("Vendor VAT No","/OPT/VIM_BL_1RIDX_OCR_DATA-VENDOR_VAT_NO","",DT_OPTVIM_7AX2_2002_VENDOR_VAT_NO,False)
Call SetTextbox("VAT Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-VAT_AMOUNT","","",False)
Call SetTextbox("Vendor Tax No","/OPT/VIM_BL_1RIDX_OCR_DATA-VENDOR_TAX_NO","",FormatBlank(DT_OPTVIM_7AX2_2002_VENDOR_TAX_NO),False)
Call SetTextbox("Tax Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-TOT_TAX_AMOUNT","","",False)
Call SetTextbox("Recepient VAT No","/OPT/VIM_BL_1RIDX_OCR_DATA-RECEPIENT_VAT_NO","",DT_OPTVIM_7AX2_2002_RECEPIENT_VAT_NO,False)
Call SetTextboxNoLabel("/OPT/VIM_BL_1RIDX_OCR_DATA-TAX_CODE",0,"",False)
Call SetTextboxNoLabel("/OPT/VIM_BL_1RIDX_OCR_DATA-TAX_RATE",0,"",False)
Call SetTextbox("Fiscal Rep VAT No","/OPT/VIM_BL_1RIDX_OCR_DATA-FISCAL_REP_VAT_N","","",False)
Call SetTextbox("Tax Rate 1","/OPT/VIM_BL_1RIDX_OCR_DATA-TAXRATE_1","","",False)
Call SetTextbox("Tax Amount 1","/OPT/VIM_BL_1RIDX_OCR_DATA-TAXAMT_1","","",False)
Call SetTextbox("Tax Rate 2","/OPT/VIM_BL_1RIDX_OCR_DATA-TAXRATE_2","","",False)
Call SetTextbox("Tax Amount 2","/OPT/VIM_BL_1RIDX_OCR_DATA-TAXAMT_2","","",False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN","","",False)
Call SetTextbox("Supply Date","/OPT/VIM_BL_1RIDX_OCR_DATA-SUPPLY_DATE","","",False)
Call SetTextbox("Payment Terms","/OPT/VIM_BL_1RIDX_OCR_DATA-PYMNT_TERMS","","",False)
Call SetTextbox("CM REF DAT","/OPT/VIM_BL_1RIDX_OCR_DATA-CM_REF_DAT","","",False)

Call SetTextbox("Bank name","/OPT/VIM_BL_1RIDX_OCR_DATA-BANKA","",DT_OPTVIM_7AX2_2002_BANK_NAME,False)
Call SetTextbox("Bank Account","/OPT/VIM_BL_1RIDX_OCR_DATA-BANKN","",DT_OPTVIM_7AX2_2002_BANK_ACCOUNT,False)
Call SetTextbox("IBAN","/OPT/VIM_BL_1RIDX_OCR_DATA-IBAN","",DT_OPTVIM_7AX2_2002_IBAN,False)
Call SetTextbox("Bank number","/OPT/VIM_BL_1RIDX_OCR_DATA-BANKL","",DT_OPTVIM_7AX2_2002_BANK_NUMBER,False)
Call SetTextbox("Doc\.Header Text","/OPT/VIM_BL_1RIDX_OCR_DATA-BKTXT","",DT_ZTPVM_DOC_DET_0100_CHECK_TEXT_OF_SWIFTBIC,False)
Call SetTextbox("Doc\.Header Text","/OPT/VIM_BL_1RIDX_OCR_DATA-BKTXT","","",False)

Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call PressEnter() 
Call ClickButton("SUBMIT   \(F8\)",False)C

Call GetTextboxValue("MESSTXT1",0,"DT_OPTVIM_7AX2_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT",True)
Call GetTextboxValue("MESSTXT2",0,"DT_OPTVIM_7AX2_0010_CHECK_TEXT_OF_MESSTXT2_OUTPUT",True)
Call VerifyTextBoxContent("Information Message","MESSTXT1",0,UCASE(DT_OPTVIM_7AX2_0010_CHECK_TEXT_OF_MESSTXT1_OUTPUT),True)
Call ClickButton("Continue   \(Enter\)",True)


Call LogOff()
Call FinalStatus ()



