

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.04.03_AdjustmentPostings_Vendors without clear_GL_Indic
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.04.01.04.03_AdjustmentPostings_Vendors without clear_GL_Indic"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.04.03_AdjustmentPostings_Vendors without clear_GL_Indic.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-F-42 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F42_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Period","BKPF-MONAT","",ConvertDoubledigit(CSTR(Month(DT_F42_0100_PERIOD))),False)
Call SetTextbox("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_F42_0100_POSTING_DATE), False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F42_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F42_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0100_PSTKY,False)
Call SetTextbox("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_F42_0100_DOCUMENT_DATE), False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F42_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F42_0100_REFERENCE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0100_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F42_0100_SGL_IND,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F42_0304_TAX_CODE,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0304_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0304_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0304_ACCOUNT,False)

Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0300_AMOUNT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_F42_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F42_0330_BSEGHKTID,False)
Call TakeScreenShot
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)


Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()

Call GetStatusBar("item1","DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02" )


Call WriteRunTimeDataToExcelGlobalSheet ("DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot


'''''--------TransactionCode-FB03 ----------''''
'
Call SetTcode(DT_F42_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F42_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F42_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",Year(DT_F42_0100_FISCAL_YEAR),False)
Call TakeScreenShot
Call PressEnter()     

DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))

Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 1, "AZBET", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "RF05A_UBAZW", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)

'Columns not available in the grid, Need to be validated after adding the columns
'Call VerifyGridCellContent("", 1, "PSWBT", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWBT)
'Call VerifyGridCellContent("", 1, "PSWSL", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL)
'Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
'Call VerifyGridCellContent("", 1, "DMBTR", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)


Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "ZUONR", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "RF05A_UBAZW", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

'Columns not available in the grid, Need to be validated after adding the columns
'Call VerifyGridCellContent("", 2, "PSWBT", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWBT)
'Call VerifyGridCellContent("", 2, "PSWSL", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)
'Call VerifyGridCellContent("", 2, "HKONT", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
'Call VerifyGridCellContent("", 2, "DMBTR", 0, DT_F42_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMBTR)

Call LogOff'
Call FinalStatus()
