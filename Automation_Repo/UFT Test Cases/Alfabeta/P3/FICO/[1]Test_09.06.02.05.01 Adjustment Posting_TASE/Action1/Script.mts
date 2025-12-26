

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.02.05.01 Adjustment Posting
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

gstrTestCaseName = "Test_09.06.02.05.01 Adjustment Posting"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
'''''''--------TransactionCode-F-21----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F21_0100_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date","BKPF-BUDAT", "", ConvertDate(DT_F21_0100_POSTING_DATE), False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F21_0100_SGL_IND,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0304_AMOUNT,False)
Call SetTextbox("Due On","BSEG-ZFBDT","",ConvertDate(DT_F21_0304_DUE_ON),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0304_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0304_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0301_TEXT,False)

Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_F21_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F21_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet("DT_F21_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_F21_0100_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode-FBL5N----------''''
'
Call SetTcode(DT_F21_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F21_1000_CUSTOMER_ACCOUNT,False)
Call SelectRadioButton("X_AISEL", "All items", False)

Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F21_1106_DOCUMENT_NUMBER,True)

Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

''Call ClickLabel(DT_F21_1106_DOCUMENT_NUMBER, 0, False)
Call DoubleClickGuiGridCell("", 0, 1, "Document Number", False)


Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 2, "Amount", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "Currency",0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "Text", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)

Call VerifyGridCellContent("", 1, "Account", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F21_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)

Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type", "BKPF-BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART, True)
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT), True)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

''''''--------TransactionCode-faglb03----------''''
'
Call SetTcode(DT_F21_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW", 0, DT_F21_1000_ACCOUNT_NUMBER, False)
Call clickButton("Choose Ledger   \(Ctrl\+F4\)",False)
Call SetTextbox("Ledger","SVALD-VALUE", 0, DT_F21_0300_LEDGER, False)

Call ClickButton("Cont\.   \(Enter\)",True)
Call ClickButton("Execute   \(F8\)",False)

Call DoubleClickGuiGridCell("", 0, Cint(DT_MONTH)+1, "Balance", False)

''Call SetFocusGuiLabel(DT_F21_1105_DOCUMENT_NUMBER, "", "", False)
Call SelectColumnGuiGrid("", 0, "Document Number", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", 0, DT_F21_1106_DOCUMENT_NUMBER, True)
Call PressEnter()  
Call TakeScreenShot
Call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Call TakeScreenShot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "Amount",0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)

Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type", "BKPF-BLART", 0, DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1, True)
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT", 0, ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1), True)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Call LogOff()
Call FinalStatus()

