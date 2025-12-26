

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p14
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

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p14"
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

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F37_0113_COMPANY_CODE,False)
Call SetTextbox("Trg.sp.G/L ind.","RF05A-ZUMSK","",DT_F37_0113_TRGSPGL_IND,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F37_0113_REFERENCE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F37_0113_CURRENCYRATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F37_0113_DOCUMENT_DATE),False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F37_0113_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F37_0113_ACCOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F37_0113_TYPE,False)
Call PressEnter() 
Call TakeScreenShot()
Call SelectCheckbox("RF05A-XMWST",0,"ON",False)
Call SetTextbox("Due On","BSEG-ZFBDT","",ConvertDate(DT_F37_0304_DUE_ON),False)
Call SetTextbox("Tax code","BSEG-MWSKZ","",DT_F37_0304_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F37_0304_AMOUNT,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_GET")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_F37_0113_CHECK_TEXT_OF_STATUSBAR)
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Company code","", DT_F37_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyGridCellContent("",1,"Account","", DT_F37_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_F37_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
