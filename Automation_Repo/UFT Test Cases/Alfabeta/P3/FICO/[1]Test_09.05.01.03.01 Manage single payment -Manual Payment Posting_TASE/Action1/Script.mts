		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.08.02 Manage Wholesale Store Customer Cheques -Reversal
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

gstrTestCaseName = "Test_09.05.01.08.02 Manage Wholesale Store Customer Cheques -Reversal"
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
''''--------TransactionCode-Ftr04----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

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
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR04_1000_COMPANY_CODE,False) 
Call TakeScreenShot
Call SelectRadioButton("PRESBANK","From ""Presented to Bank"" To:",False)
Wait(2)
Call SelectRadioButton("X3_C","Bill of Exchange Portfolio",False)
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot

Call SetTextbox("Interim Account","P_BANKN","",DT_FTR04_1002_INTERIM_ACCOUNT,False) 
Call SetTextbox("Document Number","S_BELNR-LOW","",DT_FTR04_1002_DOCUMENT_NUMBER_OCC1,False) 
Call SetTextbox("Bill of Exchange Portfolio","P_PORT1","",DT_FTR04_1002_BILL_OF_EXCHANGE_PORTFOLIO,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR04_1002_CURRENCY,False) 
Call SetTextbox("Reason for Reversal","P_STGRD","",DT_FTR04_1002_REASON_FOR_REVERSAL,False) 

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
CAll SelectCheckboxNoLabel(0,DT_FTR04_0120_NO_NAME_OCC1,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot

''''--------TransactionCode-fb03----------''''
Call SetTcode(DT_FTR04_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FTR04_0100_DOCUMENT_NUMBER,False) 
Call PressEnter()     
Call TakeScreenShot

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FTR03_0750_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FTR03_0750_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_FTR04_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)

Call LogOff'
Call FinalStatus()
