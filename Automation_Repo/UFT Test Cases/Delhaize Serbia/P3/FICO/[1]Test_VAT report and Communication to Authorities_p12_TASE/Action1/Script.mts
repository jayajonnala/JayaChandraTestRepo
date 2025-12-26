

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p12
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

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p12"
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

Call SelectRadioButton("RF05A-XPOS1","Transfer posting with clearing",False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB05_0122_DOCUMENT_DATE),False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_FB05_0122_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB05_0122_SGL_IND,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()
Call SelectCheckbox("RF05A-XMWST",0,"ON",False)
Call SetTextbox("Tax code","BSEG-MWSKZ","",DT_FB05_0304_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0304_AMOUNT,False)
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call SelectCheckbox("RF05A-XNOPS",0,"OFF",False)
Call SetTextbox("Account","RF05A-AGKON","",DT_FB05_0710_ACCOUNT,False)
Call SetTextbox("Special G/L ind","RF05A-AGUMS","",DT_FB05_0710_SPECIAL_GL_IND,False)
Call TakeScreenShot()
Call PressEnter() 
Call ClickButton("Select All",False) 
Call ClickButton("Deactivate Items",False) 
Call SelectCellGuiTable("SAPDF05XTC_6102","Document Number","Document Number",DT_FB05_0731_FROM,False)
Call ClickButtonIfExist("Activate Items",False)
''Call ClickButtonIfExist("Requests   \(F6\)",False) 
''Call TakeScreenShot()
''
''Call ClickButtonIfExist("Requests   \(F6\)",False) 
''Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
wait 3
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
wait 2
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_OP_DOCUMENT_CREATED")
GetRowNo = 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_DOCUMENT_TO_VERIFY)
Call SelectMenuBar("Document;Display")
Call TakeScreenShot
Call VerifyGridCellContent("",1,"Company code","", DT_F47_0750_CHECK_TEXT_OF_COMPANY_CODE)
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_FB05_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyGridCellContent("",1,"Account","", DT_F47_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",1,"Posting Key","", DT_F47_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
