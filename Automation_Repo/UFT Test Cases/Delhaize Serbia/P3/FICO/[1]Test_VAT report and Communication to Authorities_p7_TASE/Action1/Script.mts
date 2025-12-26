

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p7
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p7"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''--------TransactionCode-FB60----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,False)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB60_1000_COMPANY_CODE)
Call TakeScreenShot
'Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB60_0010_POSTING_DATE),False)
'Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB60_0010_TAX_AMOUNT,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
''Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
'Call SetTextbox("Amount","INVFO-WAERS","",DT_FB60_0010_AMOUNT_OCC1,False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE,False)
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_FB60_0010_REFERENCE,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_FB60_0010_INVOICE_DATE),False)
Call SelectCheckbox("INVFO-XMWST",0, "ON", False)
Call TakeScreenShot
Call PressEnter()
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "",DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "",DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area",1, "", "",DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
wait 2
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOCUMENT_CREATED")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)
'''Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE_OCC1,False)
'''Call TakeScreenShot
''''Call PressEnter()
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB60_1000_COMPANY_CODE)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Company code","", DT_FB60_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB60_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyGridCellContent("",1,"Account","", DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","", DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account","", DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","", DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"Posting Key","", DT_FB60_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

'''''--------TransactionCode-S_ALR_87012357----------''''

Call SetTcode(DT_FB60_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Get Variant...   \(Shift\+F5\)", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Created By","ENAME-LOW","","",True)
'Call SetTextbox("Variant","V-LOW","",DT_S_ALR_87012357_0600_GRIDCELL_42_VARIANT_NAME,True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RFUMSV00.*","",10, True)
Call ClickButton("Choose   \(F2\)", True)
Call ClickButtonIfExist("PUSHB_O1",False)
Call ClickButtonIfExist("PUSHB_O2",False)
''Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_FB60_1000_DOCUMENT_NUMBER,False)
Call SetTextboxNoLabel("BR_BELNR-LOW","",DT_DOCUMENT_NUM,False)
'''Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","","2021",False)
Call SetTextboxNoLabel("BR_GJAHR-LOW","",DT_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_FB60_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_FB60_1000_TO),False)
Call SelectCheckbox("PAR_DEF",0,"OFF",False)
Call SetTextbox("Tax on sales/purchases code","SEL_MWKZ-LOW","",DT_FB60_1000_TAX_ON_SALESPURCHASES_CODE,False)
Call SetTextbox("Program run date","PAR_LAUD","",ConvertDate(DT_FB60_1000_PROGRAM_RUN_DATE),False)
Call SetTextbox("Identification","PAR_LAUI","",DT_FB60_1000_IDENTIFICATION,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0120_CHECK_TEXT_OF_QG,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)
Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME_OCC2,1)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
