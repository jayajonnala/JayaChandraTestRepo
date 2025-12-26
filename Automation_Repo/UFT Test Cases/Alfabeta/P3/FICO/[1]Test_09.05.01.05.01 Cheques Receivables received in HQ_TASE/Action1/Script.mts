		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.05.01 Cheques Receivables received in HQ
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

gstrTestCaseName = "Test_09.05.01.05.01 Cheques Receivables received in HQ"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-F-36----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectRadioButton("RF05A-XPOS1","Incoming payment",False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F36_0122_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F36_0122_TYPE,False) 
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F36_0122_COMPANY_CODE,False) 
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F36_0122_POSTING_DATE),False) 
Call SetTextbox("Period","BKPF-MONAT","",DT_F36_0122_PERIOD,False) 
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F36_0122_CURRENCYRATE,False) 
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F36_0122_PSTKY,False) 
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F36_0122_SGL_IND,False) 
Call SetTextbox("Account","RF05A-NEWKO","",DT_F36_0122_ACCOUNT,False) 
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Bill Portfolio","RF05A-PORTF","",DT_F36_2100_BILL_PORTFOLIO,True)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Due On","BSEG-ZFBDT","",ConvertDate(DT_F36_0320_DUE_ON),False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F36_0320_AMOUNT,False) 
Call SetTextbox("Check No\.","BSED-BOENO","",DT_F36_0320_CHECK_NO,False)  
Call SetTextbox("Bank/Acct No","BSED-BANK","",DT_F36_0320_BANKACCT_NO,False) 
Call SetTextbox("/","BSED-ACCOU","",DT_F36_0320_BSEDACCOU,False) 
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F36_0320_PSTKY,False) 
Call SetTextbox("Account","RF05A-NEWKO","",DT_F36_0320_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F36_0301_AMOUNT,False) 
Call TakeScreenShot
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetTextboxValue("MESSTXT1",0,"DT_F36_0010_CHECK_TEXT_OF_MESSTXT1_OCC1_OUTPUT",True)
Call WriteRunTimeDataToExcelGlobalSheet("DT_F36_0010_CHECK_TEXT_OF_MESSTXT1_OCC1_OUTPUT",DT_F36_0010_CHECK_TEXT_OF_MESSTXT1_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyTextBoxContent("Information Message","MESSTXT1",0,Lcase(DT_F36_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)

''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_F36_0122_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F36_0100_DOCUMENT_NUMBER,False) 
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F36_0100_COMPANY_CODE,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F36_0100_FISCAL_YEAR,False) 
Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_ACCNT_1)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ACCNT_2)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
''
''Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
''Call VerifyGridCellContent("", 2, "KTONR", 0, DT_ACCNT_2)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
''Call VerifyGridCellContent("", 2, "Currency", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)
''Call VerifyGridCellContent("", 2, "Amount", 0, DT_F36_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call Logoff()
Call FinalStatus()

