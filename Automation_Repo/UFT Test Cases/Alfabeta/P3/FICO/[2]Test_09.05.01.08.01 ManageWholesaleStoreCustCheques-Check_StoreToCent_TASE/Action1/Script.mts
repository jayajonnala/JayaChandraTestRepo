		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.04.12 Manage Cheques Payable- Create Check Information
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.05.01.04.12 Manage Cheques Payable- Create Check Information"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-FB01----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0302_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","1",DT_FB01_0302_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_FB01_0300_PROFIT_CTR,False)  ''Added as part of OneClickMaintenance on 21/06/2024''
Call TakeScreenShot
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("House Bank","BSEG-HBKID","",DT_FB01_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_FB01_0330_BSEGHKTID,False)
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("C", "RF05A-AZSAL", 0, DT_FB01_0700_CHECK_TEXT_OF_C, False)

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FCHI----------''''
Call SetTcode(DT_FB01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("House bank","PCEC-HBKID","",DT_FB01_0100_HOUSE_BANK,False)
Call SetTextbox("Paying company code","PCEC-ZBUKR","",DT_FB01_0100_PAYING_COMPANY_CODE,False)
Call SetTextbox("Account ID","PCEC-HKTID","",DT_FB01_0100_ACCOUNT_ID,False)
Call TakeScreenShot

Call ClickButton("Display   \(Shift\+F4\)",False)
Call TakeScreenShot

Call GetTableCellData("SAPMFCHICHKI_CONTROL", "Number status", 5, "", "", "DT_FB01_0200_CHECK_TEXT_OF_TABLECELL_NUMBER_STATUS_0_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB01_0200_CHECK_TEXT_OF_TABLECELL_NUMBER_STATUS_0_OUTPUT",DT_FB01_0200_CHECK_TEXT_OF_TABLECELL_NUMBER_STATUS_0)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FCH5----------''''
Call SetTcode(DT_FB01_0200_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Payment document no\.","PAYR-VBLNR","",DT_FB01_0500_PAYMENT_DOCUMENT_NO,False)
Call SetTextbox("Paying company code","PAYR-ZBUKR","",DT_FB01_0500_PAYING_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","PAYR-GJAHR","",Year(DT_FB01_0500_FISCAL_YEAR),False)
Call SetTextbox("House bank","PAYR-HBKID","",DT_FB01_0500_HOUSE_BANK,False)
Call SetTextbox("Account ID","PAYR-HKTID","",DT_FB01_0500_ACCOUNT_ID,False)
Call SetTextbox("Check number","PAYR-CHECT","",DT_FB01_0500_CHECK_NUMBER,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call Sendkey("{F2}")
Wait 5
'Call VerifyTextBoxContent("Payee name","PAYR-ZNME1", 0, "CRETA FARMS ΑΒΕΕ", False)
Call VerifyTextBoxContent("Amount paid","PAYR-RWBTR", 0, DT_FB01_0501_CHECK_TEXT_OF_AMOUNT_PAID, False)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item4","DT_FB01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(Lcase(DT_FB01_0500_CHECK_TEXT_OF_STATUSBAR))

Call LogOff'
Call FinalStatus()
