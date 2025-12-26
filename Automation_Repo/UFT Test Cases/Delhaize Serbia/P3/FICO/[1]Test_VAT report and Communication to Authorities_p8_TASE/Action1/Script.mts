

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p8
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

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p8"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'''
''--------TransactionCode-FB75----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB75_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB75_1000_COMPANY_CODE)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",true)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB75_0510_AMOUNT,False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB75_0510_REFERENCE,False)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB75_0510_CUSTOMER,False)
Call SetTextbox("Document date","INVFO-BLDAT","",ConvertDate(DT_FB75_0510_DOCUMENT_DATE),False)
Call SelectCheckbox("INVFO-XMWST",0, "ON", False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB75_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB75_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "",DT_FB75_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area",1, "", "",DT_FB75_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center",1, "", "",DT_FB75_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB75_0550_HEADERTEXT,False)
Call TakeScreenShot
Call SelectTab("TS", "Payment",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_OP_DOCUMENT_CREATED")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB75_1000_COMPANY_CODE_OCC1,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB75_1000_COMPANY_CODE_OCC1)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",true)
'Call PressEnter()
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Company code","", DT_FB75_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB75_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyGridCellContent("",1,"Account","", DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","", DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account","", DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","", DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"Posting Key","", DT_FB75_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
