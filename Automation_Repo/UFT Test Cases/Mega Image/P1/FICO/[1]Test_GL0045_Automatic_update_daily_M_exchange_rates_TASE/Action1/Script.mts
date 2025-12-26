
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0016 Upload GL Document with vendor account on both ledger
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
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If

gstrTestCaseName = "Test_GL0045_Automatic_update_daily_M_exchange_rates_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''
''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''''----------------------Tcode ZFIGL_RFIMPECB----------------------------
'''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()   
Call TakeScreenshot()' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Date in Data Source","GP_DATE","",ConvertDAte(DT_ZFIGL_RFIMPECB_1000_DATE_IN_DATA_SOURCE),False)
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenshot()' 
Call SetTextboxPopupIfExist("ENAME-LOW", "Created By", "")
Call ClickButtonIfExist("Execute   \(F8\)",True)
''CAll SelectRowGuiGrid("Variant Catalog for Program RFIMPNBS", 0, "Variant name", DT_ZFIGL_RFIMPECB_600_GRIDCELL_1_VARIANT_NAME, True)
CAll SelectRowGuiGrid("Variant Catalog.*", 0, "Variant name", DT_ZFIGL_RFIMPECB_600_GRIDCELL_1_VARIANT_NAME, True)

Call TakeScreenshot()' 
Call ClickButton("Choose   \(F2\)",True)
Call TakeScreenshot()' 
Call SelectCheckBox("GP_TEST",0,DT_ZFIGL_RFIMPECB_1000_TEST,False)
Call TakeScreenshot()'

Call SetTextbox("Date in Data Source","GP_DATE","",ConvertDate(DT_ZFIGL_RFIMPECB_1000_DATE_IN_DATA_SOURCE),False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot()'
' Updated x, y co ordinates
Call GetLabelContentByRefLabel("EUR","-161","0","DT_EXPECTED_CURRENCY_RATE_OUTPUT",False)
Call GetLabelContentByRefLabel("GBP","-161","0","DT_EXPECTED_CURRENCY_RATE1_OUTPUT",False)
Call GetLabelContentByRefLabel("RSD","-161","0","DT_EXPECTED_CURRENCY_RATE2_OUTPUT",False)
Call GetLabelContentByRefLabel("USD","-161","0","DT_EXPECTED_CURRENCY_RATE3_OUTPUT",False)
'
'Call GetLabelContentByRefLabel("EUR","-299","0","DT_EXPECTED_CURRENCY_RATE_OUTPUT",False)
'Call GetLabelContentByRefLabel("GBP","-299","0","DT_EXPECTED_CURRENCY_RATE1_OUTPUT",False)
'Call GetLabelContentByRefLabel("RSD","-299","0","DT_EXPECTED_CURRENCY_RATE2_OUTPUT",False)
'Call GetLabelContentByRefLabel("USD","-299","0","DT_EXPECTED_CURRENCY_RATE3_OUTPUT",False)
''
''Add Get functions here
''Call VerifyifGuiLabelExists(ConvertDate(DT_ZFIGL_RFIMPECB_120_CHECK_TEXT_OF_NO_NAME))
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_ZFIGL_RFIMPECB_120_CHECK_TEXT_OF_NO_NAME),0)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenshot()'
Call SelectCheckBox("GP_TEST",0,"OFF",False)
Call TakeScreenshot()'
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot()'
Call VerifyifGuiLabelExists(DT_ZFIGL_RFIMPECB_120_CHECK_TEXT_OF_UPDATE_RUN)
''Call VerifyifGuiLabelExists(ConvertDate(DT_ZFIGL_RFIMPECB_120_CHECK_TEXT_OF_NO_NAME_OCC2))
Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_ZFIGL_RFIMPECB_120_CHECK_TEXT_OF_NO_NAME_OCC2),0)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenshot()'
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenshot()'

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'
'''''''----------------------Tcode OB08----------------------------

''Enter the Tcode
Call SetTcode(DT_ZFIGL_RFIMPECB_100_OKCD) 
Call PressEnter()   
Call TakeScreenshot()' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenshot()' 
Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDAte(DT_DESIRED_DATE),True)
Call TakeScreenshot()' 
Call ClickbuttonIfExist("Continue   \(Enter\)",True)
'
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call VerifyTableCellContent(1, "Exchange Rate Type", "SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_EXCHANGE_TYPE)
'Call VerifyTableCellContent(1, "Valid From", "SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
'Call VerifyTableCellContent(1, "From Currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
'Call VerifyTableCellContent(1, "Direct quoted exchange rate","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE_OUTPUT)
'Call VerifyTableCellContent(1, "To-currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
'CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Direct quoted exchange rate", 1, "", "", "DT_EXPECTED_CURRENCY_RATE_OUTPUT", False)

Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
Call VerifyTableCellContent(DT_ROW_EXRT,"ValidFrom","SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_FROM_CURRENCY)
''Call VerifyTableCellContent(Replace(DT_ROW_EXRT,",",""),"Dir.quot.","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE)
Call VerifyTableCellContent(Replace(DT_ROW_EXRT,",",""),"Dir.quot.","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_TO_CURRENCY)
CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", DT_ROW_EXRT, "", "", "DT_RATE_OUTPUT", False)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenshot()' 
Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY1,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDAte(DT_DESIRED_DATE),True)
Call TakeScreenshot()' 
Call Clickbutton("Continue   \(Enter\)",True)
Call TakeScreenshot()' 
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Call VerifyTableCellContent(1, "Exchange Rate Type", "SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1, "Valid From", "SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
''Call VerifyTableCellContent(1, "From Currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
''CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Direct quoted exchange rate", 1, "", "", "DT_EXPECTED_CURRENCY_RATE1_OUTPUT", False)
''Call VerifyTableCellContent(1, "Direct quoted exchange rate","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE1_OUTPUT)
''Call VerifyTableCellContent(1, "To-currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)

Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(DT_ROW_EXRT,"ValidFrom","SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_FROM_CURRENCY1)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE1)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_TO_CURRENCY)
CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", DT_ROW_EXRT, "", "", "DT_RATE1_OUTPUT", False)


Call ClickButton("Position\.\.\.",False)
Call TakeScreenshot()' 
Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY2,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDAte(DT_DESIRED_DATE),True)
Call TakeScreenshot()' 
Call Clickbutton("Continue   \(Enter\)",True)
Call TakeScreenshot()' 

Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Call VerifyTableCellContent(1, "Exchange Rate Type", "SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1, "Valid From", "SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
''Call VerifyTableCellContent(1, "From Currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
''CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Direct quoted exchange rate", 1, "", "", "DT_EXPECTED_CURRENCY_RATE2_OUTPUT", False)
''Call VerifyTableCellContent(1, "Direct quoted exchange rate","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE2_OUTPUT)
''Call VerifyTableCellContent(1, "To-currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)

Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call VerifyTableCellContent(DT_ROW_EXRT,"ValidFrom","SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_FROM_CURRENCY2)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE2)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_TO_CURRENCY)
CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", DT_ROW_EXRT, "", "", "DT_RATE2_OUTPUT", False)

Call ClickButton("Position\.\.\.",False)
Call TakeScreenshot()' 
Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY2,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",ConvertDAte(DT_DESIRED_DATE),True)
Call TakeScreenshot()' 
Call Clickbutton("Continue   \(Enter\)",True)
Call TakeScreenshot()' 
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Call VerifyTableCellContent(1, "Exchange Rate Type", "SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1, "Valid From", "SAPL0SAPTCTRL_V_TCURR", ConvertDate(DT_EXPECTED_VALID_DATE))
''Call VerifyTableCellContent(1, "From Currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1, "Direct quoted exchange rate","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE3_OUTPUT)
''Call VerifyTableCellContent(1, "To-currency","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
''CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Direct quoted exchange rate", 1, "", "", "DT_EXPECTED_CURRENCY_RATE3_OUTPUT", False)

Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(DT_ROW_EXRT,"ValidFrom","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_VALID_DATE)
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_FROM_CURRENCY3)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_CURRENCY_RATE3)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR", DT_EXPECTED_TO_CURRENCY)
CAll GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", DT_ROW_EXRT, "", "", "DT_RATE3_OUTPUT", False)

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

