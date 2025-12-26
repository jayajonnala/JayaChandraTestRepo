

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p11
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

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p11"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''--------TransactionCode-FB48----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F48_0110_COMPANY_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F48_0110_AMOUNT,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F48_0110_REFERENCE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F48_0110_CURRENCYRATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F48_0110_DOCUMENT_DATE),False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F48_0110_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F48_0110_ACCOUNT,False)
Call SetTextbox("Special G/L ind","RF05A-UMSKZ","",DT_F48_0110_SPECIAL_GL_IND,False)
Call SetTextbox("Account","RF05A-KONTO","",DT_F48_0110_ACCOUNT_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F48_0110_TYPE,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_F48_0110_VALUE_DATE),False)
Call TakeScreenShot()
Call ClickButtonIfExist("Requests   \(F6\)",False) 
Call TakeScreenShot()
Call ClickButtonIfExist("Descending   \(Shift\+F6\)",False) 
Call SelectRowGuiTableByRow("SAPMF05ATC_1702",1,False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_OP_DOCUMENT_CREATED")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Company code","", DT_F48_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_F48_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyGridCellContent("",1,"Account","", DT_F48_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","", DT_F48_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_F48_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","", DT_F48_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",1,"Special G/L ind.","", "")
Call VerifyGridCellContent("",2,"Special G/L ind.","",DT_F48_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
