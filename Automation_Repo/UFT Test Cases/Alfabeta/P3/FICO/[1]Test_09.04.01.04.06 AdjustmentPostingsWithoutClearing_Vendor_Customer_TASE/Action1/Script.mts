

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.04.06 AdjustmentPostingsWithoutClearing_Vendor_Customer
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

gstrTestCaseName = "Test_09.04.01.04.06 AdjustmentPostingsWithoutClearing_Vendor_Customer"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.04.06 AdjustmentPostingsWithoutClearing_Vendor_Customer.xls"
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
'
''''--------TransactionCode-F-42 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F42_0100_DOCHEADER_TEXT,False)
'Call SetTextbox("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_F42_0100_POSTING_DATE), False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F42_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F42_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0100_PSTKY,False)
Call SetTextbox("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_F42_0100_DOCUMENT_DATE), False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F42_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F42_0100_REFERENCE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SelectCheckbox("BKPF-XMWST", 0, "ON", False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F42_0302_PSTKY,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0302_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F42_0302_TAX_CODE,False)
Call TakeScreenShot
Call SetTextbox("Text","BSEG-SGTXT","",DT_F42_0302_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F42_0302_ACCOUNT,False)
Call ClickButtonIfExist("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot

Call SelectCheckbox("BSEG-XNEGP", 0, "ON", False)
Call TakeScreenShot
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F42_0301_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F42_0301_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F42_0301_TEXT,False)

Call SelectMenuBar("Document;Simulate")
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02" )

Call WriteRunTimeDataToExcelGlobalSheet ("DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F42_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
'
'
''''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_F42_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_F42_1000_VENDOR_ACCOUNT,False)
Call SetTextbox("Company Code","KD_BUKRS-LOW","",DT_F42_1000_COMPANY_CODE,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F42_1106_DOCUMENT_NUMBER,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)& ConvertDoubledigit(CSTR(Month(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))& ConvertDoubledigit(CSTR(Day(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)))


Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call TakeScreenShot

''''''--------TransactionCode-FBL5N ----------''''

Call SetTcode(DT_F42_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F42_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F42_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Layout","PA_VARI","",DT_FBL5N_LAYOUT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)

Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F42_1106_DOCUMENT_NUMBER_OCC1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot


'DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP = year(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)& ConvertDoubledigit(CSTR(Month(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)))& ConvertDoubledigit(CSTR(Day(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)))


'Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_IP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER_OCC1)
'Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)
'Call VerifyGridCellContent("", 1, "HKONT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
'Call VerifyGridCellContent("", 1, "BUDAT", 0, DT_F42_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT)
Call TakeScreenShot
Call LogOff'
Call FinalStatus()
