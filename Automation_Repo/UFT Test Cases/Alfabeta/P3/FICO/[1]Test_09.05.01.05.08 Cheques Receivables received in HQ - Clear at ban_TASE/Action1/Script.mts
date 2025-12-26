		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.05.08 Cheques Receivables received in HQ - Clear at ban
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

gstrTestCaseName = "Test_09.05.01.05.08 Cheques Receivables received in HQ - Clear at ban"
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

'''''--------TransactionCode-F-36----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectRadioButton("X1","Presented to Bank",False)
Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_FTR03_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_FTR03_1000_POSTING_DATE),False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR03_1000_COMPANY_CODE,False) 
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot
Call SetTextbox("Bill of Exchange/Check Portf\.","P_PORT1","",DT_FTR03_1002_BILL_OF_EXCHANGECHECK_PORTF,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call ClickLabel("DocumentNo", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FTR03_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call VerifyifGuiLabelExistsByRelativeid(DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[71,7\]")
CAll SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME,false)
Call TakeScreenshot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot
Call SetTextbox("Interim Account","RFIDTRBOE-BANKN","",DT_FTR03_1004_INTERIM_ACCOUNT,True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot
Call ClickButton("Back   \(F3\)",False)
Wait 2
Call TakeScreenshot
Call SelectRadioButton("X3","Cleared at Bank",False)
Call TakeScreenshot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot
Call SetTextbox("Interim Account","P_BANKN","",DT_FTR03_1002_INTERIM_ACCOUNT,False) 
Call SetTextbox("Bill of Exchange Portfolio","S_PORT1-LOW","",DT_FTR03_1002_BILL_OF_EXCHANGE_PORTFOLIO,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY_OCC1,False) 
Call TakeScreenshot
Call ClickButton("Execute   \(F8\)",False)

Call ClickLabel("DocumentNo", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FTR03_1105_DOCUMENT_NUMBER_OCC1,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call VerifyifGuiLabelExistsByRelativeid(DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[82,7\]")
CAll SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME_OCC1,false)
Call TakeScreenshot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot
Call SetTextbox("Value Date","RFIDTRBOE-VALUT","",ConvertDate(DT_FTR03_1007_VALUE_DATE),True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot

''''''--------TransactionCode-ftr02----------''''
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
Call ClickLabel(DT_FTR03_1000_DOCUMENT_NUMBER,1,False)
Call TakeScreenShot

' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
Call GetTextboxValue("BKPF-BELNR",0,"DT_DOC1_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOC1_OUTPUT",DT_DOC1_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)

''Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
''Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
''Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
''Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call ClickButton("Back   \(F3\)",False)
Wait 2

Call ClickLabel("4",1,False)
Call TakeScreenshot
Call GetLabelContentByRefLabel("DocumentNo",0,-96,"DT_DOC2_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOC2_OUTPUT",DT_DOC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickLabel(DT_DOC2,1,False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC1)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC1)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW_OCC1)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC1)

''Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
''Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)
''Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC1)
''Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC1)

Call ClickButton("Back   \(F3\)",False)
Wait 2

Call ClickLabel("4",1,False)
Call TakeScreenshot
Call GetLabelContentByRefLabel("DocumentNo",0,-144,"DT_DOC3_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOC3_OUTPUT",DT_DOC3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickLabel(DT_DOC3,1,False)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC2)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC2)

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ_OCC1)
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC2)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC2)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC2)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC2)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC2)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC2)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)

''Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC2)
''Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC2)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC2)
''Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC2)
''Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC2)

Call ClickButton("Back   \(F3\)",False)
Wait 2

Call GetLabelContentByRefLabel("DocumentNo",0,-160,"DT_DOC4_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOC4_OUTPUT",DT_DOC4)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickLabel(DT_DOC4,1,False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC3)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC3)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC3)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC3)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC3)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC3)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OCC3)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC3)

Call VerifyGridCellContent("", 1, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW_OCC2)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW_OCC2)

''Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC3)
''Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC3)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC3)
''Call VerifyGridCellContent("", 2, "Currency", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL)
''Call VerifyGridCellContent("", 2, "Amount", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OCC3)

Call Logoff()
Call FinalStatus()

