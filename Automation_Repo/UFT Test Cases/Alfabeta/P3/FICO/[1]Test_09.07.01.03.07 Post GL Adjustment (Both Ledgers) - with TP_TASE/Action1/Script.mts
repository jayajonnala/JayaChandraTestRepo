		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.03.07 Post GL Adjustment (Both Ledgers) - with TP
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

gstrTestCaseName = "Test_09.07.01.03.07 Post GL Adjustment (Both Ledgers) - with TP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''''--------TransactionCode-S_PL0_86000030----------''''
'''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Currency Type","U3_CURTP-LOW","",DT_S_PL0_86000030_1000_CURRENCY_TYPE,False)
Call SetTextbox("Company Code","U3_RBUKR-LOW","",DT_S_PL0_86000030_1000_COMPANY_CODE,False)
Call SetTextbox("From period","PAR_01","",ConvertDoubleDigit(Cstr(MOnth(DT_S_PL0_86000030_1000_FROM_PERIOD))),False)
Call SetTextbox("To period","PAR_02","",ConvertDoubleDigit(Cstr(MOnth(DT_S_PL0_86000030_1000_TO_PERIOD))),False)
Call SetTextbox("Fiscal year","PAR_06","",Year(DT_S_PL0_86000030_1000_FISCAL_YEAR),False)
Call SetTextbox("Ledger","PAR_03","",DT_S_PL0_86000030_1000_LEDGER,False)
Call SetTextbox("Trading partner","U3_00032-LOW","",DT_S_PL0_86000030_1000_TRADING_PARTNER,False)
Call SetTextbox("Account Number","U3_00011-LOW","",DT_S_PL0_86000030_1000_ACCOUNT_NUMBER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonifExist("Adapt report width   \(Enter\)",False)

'''''''''--------TransactionCode-FB50----------''''
''
Call SetTcode(DT_S_PL0_86000030_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_S_PL0_86000030_1000_COMPANY_CODE_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document Date","ACGL_HEAD-BLDAT", "", ConvertDate(DT_S_PL0_86000030_1010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date","ACGL_HEAD-BUDAT", "", ConvertDate(DT_S_PL0_86000030_1010_POSTING_DATE), False)
Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_S_PL0_86000030_1010_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_S_PL0_86000030_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Document type","ACGL_HEAD-BLART", "", DT_S_PL0_86000030_1010_DOCUMENT_TYPE, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_S_PL0_86000030_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "G/L acct", "2", "", "", DT_S_PL0_86000030_0100_TABLECELL_GL_ACCT_1, False)

Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_S_PL0_86000030_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "2", "", "", DT_S_PL0_86000030_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)

Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_S_PL0_86000030_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "2", "", "", DT_S_PL0_86000030_0100_TABLECELL_COST_CENTER_1, False)

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "Cost center", DT_S_PL0_86000030_0100_TABLECELL_COST_CENTER_0, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{DOWN}")
Call SendKey("{TAB}")

Call SelectCellGuiTable("SAPLFSKBTABLE", "D/C", "Cost center", DT_S_PL0_86000030_0100_TABLECELL_COST_CENTER_1, False)
Call SendKey("{F4}")
Wait 2
Call SendKey("{TAB}")

Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_S_PL0_86000030_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "2", "", "", DT_S_PL0_86000030_0100_TABLECELL_TEXT_1, False)

Call SetTableData("SAPLFSKBTABLE", "Trading partner", "1", "", "",DT_S_PL0_86000030_0100_TABLECELL_TRADING_PARTNER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Trading partner", "2", "", "",DT_S_PL0_86000030_0100_TABLECELL_TRADING_PARTNER_1, False)

Call SetTableData("SAPLFSKBTABLE", "Profit center", "1", "", "", DT_S_PL0_86000030_0100_TABLECELL_PROFIT_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Profit center", "2", "", "", DT_S_PL0_86000030_0100_TABLECELL_PROFIT_CENTER_1, False)

Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_S_PL0_86000030_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "2", "", "", DT_S_PL0_86000030_0100_TABLECELL_BUSINESS_AREA_1, False)

Call TakeScreenShot
Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call PressEnter
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC6)
Call PressEnter
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC7)

Call GetStatusBar("item1","DT_S_PL0_86000030_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_S_PL0_86000030_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_PL0_86000030_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_S_PL0_86000030_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

'''''''''--------TransactionCode-S_PL0_86000030----------''''
'''
Call SetTcode(DT_S_PL0_86000030_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("LISTE","Classic drilldown report",False)

Call SetTextbox("Currency Type","U3_CURTP-LOW","",DT_S_PL0_86000030_1000_CURRENCY_TYPE_OCC1,False)
Call SetTextbox("Company Code","U3_RBUKR-LOW","",DT_S_PL0_86000030_1000_COMPANY_CODE_OCC2,False)
Call SetTextbox("Business Area","U3_00012-LOW","","",False)
Call SetTextbox("From period","PAR_01","",ConvertDoubleDigit(Cstr(MOnth(DT_S_PL0_86000030_1000_FROM_PERIOD_OCC1))),False)
Call SetTextbox("To period","PAR_02","",ConvertDoubleDigit(Cstr(MOnth(DT_S_PL0_86000030_1000_TO_PERIOD_OCC1))),False)
Call SetTextbox("Trading partner","U3_00032-LOW","",DT_S_PL0_86000030_1000_TRADING_PARTNER_OCC1,False)
Call SetTextbox("Account Number","U3_00011-LOW","",DT_S_PL0_86000030_1000_ACCOUNT_NUMBER_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyifGuiLabelExistsByRelativeid(DT_S_PL0_86000030_0120_CHECK_TEXT_OF_NO_NAME_OCC5, "wnd\[0\]/usr/lbl\[172,14\]")
Call ClickButtonifExist("Adapt report width   \(Enter\)",False)

Call LogOff()
Call FinalStatus()

