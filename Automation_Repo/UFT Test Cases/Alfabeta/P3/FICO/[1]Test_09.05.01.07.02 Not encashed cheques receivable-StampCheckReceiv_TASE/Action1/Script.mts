		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.07.02 Not encashed cheques receivable-StampCheckReceiv
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

gstrTestCaseName = "Test_09.05.01.07.02 Not encashed cheques receivable-StampCheckReceiv"
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
''''--------TransactionCode-FTR03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR03_1000_COMPANY_CODE,False) 
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY,False) 
Call SetTextbox("Bill of Exchange/Check Portf\.","P_PORT1","",DT_FTR03_1002_BILL_OF_EXCHANGECHECK_PORTF,False)
Call SetTextbox("Bill of Exchange/Check Number","S_BOENO-LOW","",DT_FTR03_3010_TABLECELL_SINGLE_VALUE_0,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME,false)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SetTextbox("Interim Account","RFIDTRBOE-BANKN","",DT_FTR03_1004_INTERIM_ACCOUNT,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call SelectRadioButton("X5","Bounced at Bank",False) 
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot
Call SetTextbox("Bill of Exchange/Check Number","S_BOENO-LOW","",DT_FTR03_1002_BILL_OF_EXCHANGECHECK_NUMBER,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME_OCC1,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SetTextbox("Portfolio for bounced bills","RFIDTRBOE-PORT5","",DT_FTR03_1008_PORTFOLIO_FOR_BOUNCED_BILLS,True)
Call SetTextbox("Special G/L indicator","RFIDTRBOE-UMSKZ","",DT_FTR03_1008_SPECIAL_GL_INDICATOR,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot

Call GetLabelContentByRefLabel("DocumentNo",0,-32,"DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_FTR03_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS,False) 
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FTR03_0100_DOCUMENT_NUMBER,False) 
Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

Call LogOff'
Call FinalStatus()
