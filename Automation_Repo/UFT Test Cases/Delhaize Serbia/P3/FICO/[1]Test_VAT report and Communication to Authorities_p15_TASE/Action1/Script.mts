

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p15
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

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p15"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''--------TransactionCode-F-29----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F29_0111_COMPANY_CODE,False)
Call SetTextbox("Special G/L ind","RF05A-UMSKZ","",DT_F29_0111_SPECIAL_GL_IND,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F29_0111_REFERENCE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F29_0111_CURRENCYRATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F29_0111_DOCUMENT_DATE),False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDate(DT_F29_0111_VALUE_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F29_0111_POSTING_DATE),False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F29_0111_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F29_0111_ACCOUNT,False)
Call SetTextbox("Account","RF05A-KONTO","",DT_F29_0111_ACCOUNT_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F29_0111_TYPE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F29_0111_AMOUNT,False)
Call SetTextbox("Period","BKPF-MONAT","",Month(date),False)
Call TakeScreenShot()
Call ClickButtonIfExist("Requests   \(F6\)",False)
Call TakeScreenShot()

'''Call SelectRowGuiTableByRow("SAPMF05ATC_1702",1,False)

Call SelectRowGuiTable("SAPMF05ATC_1702","Document Number",DT_DOCUMENT,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
wait 2
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
wai 2
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_F29_0111_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_F29_0111_CHECK_TEXT_OF_STATUSBAR)
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyTextBoxContent("Document Number","BKPF-BELNR",0,DT_F29_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyGridCellContent("",1,"Company code","", DT_F29_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyGridCellContent("",2,"Account","", DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SAKNR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","", DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
