
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0046_Maintain_Exhange_Rates_type_F_TASE
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


gstrTestCaseName = "Test_GL0046_Maintain_Exhange_Rates_type_F_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''----------------------Tcode TCURMNT----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot


CAll  FindRowNumber("SAPMCURRTC_WORKSETS","Worklist",DT_WORKLIST_NAME,"ROW_NO_OUTPUT")
'Call WriteRunTimeDataToExcelGlobalSheet ("ROW_NO_OUTPUT",ROW_NO)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

CAll SelectRowGuiTableByRow("SAPMCURRTC_WORKSETS", ROW_NO_OUTPUT, False)
Call ClickButton("Enter Exchange Rates for Worklist   \(F2\)",False)

Call SetTableData("SAPMCURRTC_RATE_SHEET","Valid from",1,"","",convertdate(DT_TCURMNT_200_VALID_FROM),False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Valid from",2,"","",convertdate(DT_TCURMNT_200_VALID_FROM),False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Valid from",3,"","",convertdate(DT_TCURMNT_200_VALID_FROM),False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Valid from",4,"","",convertdate(DT_TCURMNT_200_VALID_FROM),False)

Call SetTableData("SAPMCURRTC_RATE_SHEET","Exch. Rate",1,"","",DT_EXCH_RATE_EUR,False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Exch. Rate",2,"","",DT_EXCH_RATE_GBP,False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Exch. Rate",3,"","",DT_EXCH_RATE_RSD,False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Exch. Rate",4,"","",DT_EXCH_RATE_USD,False)

Call SetTableData("SAPMCURRTC_RATE_SHEET","Rate 1:1",1,"","",DT_EXCH_RATE_EUR,False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Rate 1:1",2,"","",DT_EXCH_RATE_GBP,False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Rate 1:1",3,"","",DT_EXCH_RATE_RSD,False)
Call SetTableData("SAPMCURRTC_RATE_SHEET","Rate 1:1",4,"","",DT_EXCH_RATE_USD,False)
Call TakeScreenShot

'Call ClickCellTable("SAPMCURRTC_RATE_SHEET", "Rate 1:1", 4, "", "", False)
''CAll SelectRowGuiTableByRow("SAPMCURRTC_RATE_SHEET", 4, False)

Call ClickButton("Set Worklist to ""Completed""   \(F8\)",False)
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call VerifyStatusBar(DT_TCURMNT_100_CHECK_TEXT_OF_STATUSBAR)

''Call PressEnter()
''Call PressEnter()
''Call PressEnter()
''Call PressEnter()
''Call PressEnter()
''Call PressEnter()
''Call PressEnter()


Call ClickButton("Exit   \(Shift\+F3\)",False)
''''----------------------Tcode OB08----------------------------
Call SetTcode(DT_TCURMNT_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call PressEnter()

Call ClickButton("Position\.\.\.",False)

Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",convertdate(DT_DESIRED_DATE),True)

Call ClickButton("Continue   \(Enter\)",False)
Call TakeScreenShot


Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyTableCellContent(1,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
'Call VerifyTableCellContent(1,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
'Call VerifyTableCellContent(1,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY)
'Call VerifyTableCellContent(1,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE)
'Call VerifyTableCellContent(1,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call VerifyTableCellContent(DT_ROW_EXRT,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call TakeScreenShot

Call ClickButton("Position\.\.\.",False)

Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY_OCC1,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",convertdate(DT_DESIRED_DATE),True)

Call ClickButton("Continue   \(Enter\)",False)
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call VerifyTableCellContent(1,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
''Call VerifyTableCellContent(1,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY_OCC1)
''Call VerifyTableCellContent(1,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE_OCC1)
''Call VerifyTableCellContent(1,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call VerifyTableCellContent(DT_ROW_EXRT,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY_OCC1)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE_OCC1)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call TakeScreenShot

Call ClickButton("Position\.\.\.",False)

Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY_OCC2,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",convertdate(DT_DESIRED_DATE),True)

Call ClickButton("Continue   \(Enter\)",False)
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call VerifyTableCellContent(1,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
''Call VerifyTableCellContent(1,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY_OCC2)
''Call VerifyTableCellContent(1,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE_OCC2)
''Call VerifyTableCellContent(1,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call VerifyTableCellContent(DT_ROW_EXRT,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY_OCC2)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE_OCC2)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call TakeScreenShot

Call ClickButton("Position\.\.\.",False)

Call SetTextbox("Exch\. Rate Type","SVALD-VALUE","",DT_DESIRED_EXCHANGE_TYPE,True)
Call SetTextbox("From currency","SVALD-VALUE","",DT_DESIRED_FROM_CURRENCY_OCC3,True)
Call SetTextbox("To-currency","SVALD-VALUE","",DT_DESIRED_TO_CURRENCY,True)
Call SetTextbox("Valid from","SVALD-VALUE","",convertdate(DT_DESIRED_DATE),True)

Call ClickButton("Continue   \(Enter\)",False)
Call GetTextBoxvalue("VIM_POSITION_INFO",0,"DT_ROW_OUTPUT_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call VerifyTableCellContent(1,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
''Call VerifyTableCellContent(1,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
''Call VerifyTableCellContent(1,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY_OCC3)
''Call VerifyTableCellContent(1,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE_OCC3)
''Call VerifyTableCellContent(1,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call VerifyTableCellContent(DT_ROW_EXRT,"ExRt","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_EXCHANGE_TYPE)
Call VerifyTableCellContent(DT_ROW_EXRT,"Validfrom","SAPL0SAPTCTRL_V_TCURR",convertdate(DT_EXPECTED_VALID_DATE))
Call VerifyTableCellContent(DT_ROW_EXRT,"From","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_FROM_CURRENCY_OCC3)
Call VerifyTableCellContent(DT_ROW_EXRT,"Dir.quot.","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_CURRENCY_RATE_OCC3)
Call VerifyTableCellContent(DT_ROW_EXRT,"To","SAPL0SAPTCTRL_V_TCURR",DT_EXPECTED_TO_CURRENCY)
Call TakeScreenShot



'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

