		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.06.03 Cheques Receivables received in Store - Reversal
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

gstrTestCaseName = "Test_09.05.01.06.03 Cheques Receivables received in Store - Reversal"
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

''''--------TransactionCode-Ftr04----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_FTR04_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_FTR04_1000_POSTING_DATE),False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR04_1000_COMPANY_CODE,False) 
Call TakeScreenShot

Call SelectRadioButton("CLEARBAN","From ""Cleared at Bank"" To:",False)
Wait(2)
Call SelectRadioButton("X8_C","Presented to Bank",False)
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot

Call SetTextbox("Bank Account Number","P_BANKO","",DT_FTR04_1002_BANK_ACCOUNT_NUMBER,False) 
Call SetTextbox("Bill of Exchange Portfolio","P_PORT1","",DT_FTR04_1002_BILL_OF_EXCHANGE_PORTFOLIO,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR04_1002_CURRENCY,False) 
Call SetTextbox("Reason for Reversal","P_STGRD","",DT_FTR04_1002_REASON_FOR_REVERSAL,False) 
Call SetTextbox("Bill of Exchange Number","S_BOENO-LOW","",DT_FTR04_1002_BILL_OF_EXCHANGE_NUMBER,False) 
Call SetTextbox("Document Number","S_BELNR-LOW","",DT_FTR04_1002_DOCUMENT_NUMBER,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
CAll SelectCheckboxNoLabel(0,DT_FTR04_0120_NO_NAME,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot
Call VerifyStatusBar(Lcase(DT_FTR04_0120_CHECK_TEXT_OF_STATUSBAR))

''''--------TransactionCode-fb03----------''''
Call SetTcode(DT_FTR04_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FTR04_0100_COMPANY_CODE,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FTR04_0100_FISCAL_YEAR,False) 
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)

Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call GetTextboxValue("BKPF-BELNR",0,"DT_FTR04_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR04_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",DT_FTR04_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''--------TransactionCode-Ftr04----------''''
Call SetTcode(DT_FTR04_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_FTR04_1000_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_FTR04_1000_POSTING_DATE_OCC1),False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR04_1000_COMPANY_CODE_OCC1,False) 
Call TakeScreenShot

Call SelectRadioButton("PRESBANK","From ""Presented to Bank"" To:",False)
Wait(2)
Call SelectRadioButton("X3_C","Bill of Exchange Portfolio",False)
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot

Call SetTextbox("Interim Account","P_BANKN","",DT_FTR04_1002_INTERIM_ACCOUNT,False) 
Call SetTextbox("Bill of Exchange Portfolio","P_PORT1","",DT_FTR04_1002_BILL_OF_EXCHANGE_PORTFOLIO_OCC1,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR04_1002_CURRENCY_OCC1,False) 
Call SetTextbox("Reason for Reversal","P_STGRD","",DT_FTR04_1002_REASON_FOR_REVERSAL_OCC1,False) 
Call SetTextbox("Bill of Exchange Number","S_BOENO-LOW","",DT_FTR04_1002_BILL_OF_EXCHANGE_NUMBER_OCC1,False) 
Call SetTextbox("Document Number","S_BELNR-LOW","",DT_FTR04_1002_DOCUMENT_NUMBER_OCC1,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
CAll SelectCheckboxNoLabel(0,DT_FTR04_0120_NO_NAME_OCC1,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot
Call VerifyStatusBar(Lcase(DT_FTR04_0120_CHECK_TEXT_OF_STATUSBAR_OCC1))

''''--------TransactionCode-fb03----------''''
Call SetTcode(DT_FTR04_0120_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FTR04_0100_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FTR04_0100_FISCAL_YEAR_OCC1,False) 
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ_OCC1)
Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC1)

Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ_OCC1)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)

Call GetTextboxValue("BKPF-BELNR",0,"DT_FTR04_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR04_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",DT_FTR04_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call LogOff'
Call FinalStatus()'
