		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.06.02 Cheques Receivables received in Store - Clear at
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

gstrTestCaseName = "Test_09.05.01.06.02 Cheques Receivables received in Store - Clear at"
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

'''''--------TransactionCode-FTR03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_FTR03_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_FTR03_1000_POSTING_DATE),False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR03_1000_COMPANY_CODE,False) 
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot
Call SetTextbox("Bill of Exchange/Check Portf\.","P_PORT1","",DT_FTR03_1002_BILL_OF_EXCHANGECHECK_PORTF,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY,False) 
Call SetTextbox("Bill of Exchange/Check Number","S_BOENO-LOW","",DT_FTR03_1002_BILL_OF_EXCHANGECHECK_NUMBER,False) 
Call SetTextbox("Document Number","S_BELNR-LOW","",DT_FTR03_1002_DOCUMENT_NUMBER,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

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

Call SelectRadioButton("X3","Cleared at Bank",False)
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot


Call SetTextbox("Interim Account","P_BANKN","",DT_FTR03_1002_INTERIM_ACCOUNT	,False) 
Call SetTextbox("Bill of Exchange Portfolio","S_PORT1-LOW","",DT_FTR03_1002_BILL_OF_EXCHANGE_PORTFOLIO,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY_OCC1,False) 
Call SetTextbox("Bill of Exchange/Check Number","S_BOENO-LOW","",DT_FTR03_1002_BILL_OF_EXCHANGECHECK_NUMBER_OCC1,False) 
Call SetTextbox("Document Number","S_BELNR-LOW","",DT_FTR03_1002_DOCUMENT_NUMBER_OCC1,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME_OCC1,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SetTextbox("Value Date","RFIDTRBOE-VALUT","",ConvertDate(DT_FTR03_1007_VALUE_DATE),True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot

Call VerifyStatusBar(Lcase(DT_FTR03_0120_CHECK_TEXT_OF_STATUSBAR))


'''''''--------TransactionCode-ftr02----------''''
Call SetTcode(DT_FTR03_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("APOST",0,DT_FTR03_1000_CLEARED_BILLS_OF_EXCHANGE,False)

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR03_1000_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Bill of Exchange Portfolio","S_PORTF-LOW","",DT_FTR03_1000_BILL_OF_EXCHANGE_PORTFOLIO,False) 
Call SetTextbox("Open Items at Key Date","P_STIDA","",ConvertDate(DT_FTR03_1000_OPEN_ITEMS_AT_KEY_DATE),False) 
Call SetTextbox("Check/Bill of Exchange Number","BOENO-LOW","",DT_FTR03_1000_CHECKBILL_OF_EXCHANGE_NUMBER,False) 
Call SetTextbox("Document Number","BELEG_NR-LOW","",DT_FTR03_1000_DOCUMENT_NUMBER,False) 
Call TakeScreenshot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot
Call SetFocusGuiLabel(DT_FTR03_1000_DOCUMENT_NUMBER,11,136,False)
Call TakeScreenshot
Call ClickButton("History   \(Ctrl\+Shift\+F6\)",False)
Call TakeScreenshot

Call ClickLabel("4",1,False)
Call TakeScreenshot
Call ClickLabel("4",1,False)
Call TakeScreenshot
Call ClickLabel("4",1,False)
Call TakeScreenshot
Call ClickLabel("4",1,False)
Call TakeScreenshot
Call ClickLabel("4",1,False)
Call TakeScreenshot
''Call ClickLabel("4",1,False)
''Call TakeScreenshot
''Call ClickLabel("4",1,False)
''Call TakeScreenshot

Call VerifyifGuiLabelExistsByRelativeid(DT_FTR03_0120_CHECK_TEXT_OF_PRESENTED_TO_BANK,"wnd\[0\]/usr/lbl\[22,16\]")

Call GetLabelContentByRefLabel("DocumentNo",0,-272,"DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickLabel(DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,1,False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "UMSKZ", 0, "")
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

''Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call ClickButton("Back   \(F3\)",False)
Wait 2

'Call VerifyifGuiLabelExistsByRelativeid(DT_FTR03_0120_CHECK_TEXT_OF_CLEARED_AT_BANK,"wnd\[0\]/usr/lbl\[22,16\]")

Call GetLabelContentByRefLabel("DocumentNo",0,-320,"DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickLabel(DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1,1,False)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "UMSKZ", 0, "")
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC1)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)

''Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ_OCC1)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)


Call ClickButton("Back   \(F3\)",False)
Wait 2

Call GetLabelContentByRefLabel("DocumentNo",0,-336,"DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC2_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC2_OUTPUT",DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickLabel(DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC2,1,False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "UMSKZ", 0, "")
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC2)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC2)

''Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ_OCC2)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC2)

Call Logoff()
Call FinalStatus()

