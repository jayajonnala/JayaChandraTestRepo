'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0010 Manual posting of CN to be received- 3rd party vendor direct foreign 
'.................Author : TCS      
'................ Creation Date   
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


gstrTestCaseName = "Test_CL0010 Manual posting of CN to be received- 3rd party vendor direct foreign"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0010.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-ZFIGL_UPLOAD_POST---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call RefreshExcelSheet(DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME)


Call SetTextbox("Session","P_SESS", "", DT_ZFIGL_UPLOAD_POST_1000_SESSION, False)
Call SetTextbox("File path name","P_FILE", "", DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME, False)
wait 5
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", False)

Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
If ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").exist(0) Then
	ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
Call TakeScreenShot()
ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
Wait(2)
End If

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_ZFIGL_UPLOAD_POST_500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call TakeScreenShot()						
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call ClickButtonIfExist("Yes", True)

''''--------\nSM35 --------------

Call SetTcode(DT_ZFIGL_UPLOAD_POST_120_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Sess\.:","D0100-MAPN", "", "", False)
Call TakeScreenShot
Call PressEnter()     

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButtonIfExist("Process session   \(F8\)", False)
Call SelectRadioButtonIfExist("D0300-ERROR", 0, True)
Call ClickButtonIfExist("Process   \(Enter\)", True)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Log   \(F7\)", False)
Call ClickButtonIfExist("Display   \(F2\)", False)
wait 1

Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message","", "Message Number", "312", "DT_ZFIGL_UPLOAD_POST_1400_GET_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT", False)
Call TakeScreenShot

'''''-----/nFB03 screen--------------

Call SetTcode(DT_ZFIGL_UPLOAD_POST_1400_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call ClickButtonIfExist("Display Document Header   \(F5\)", False)

Call VerifyTextBoxContent("Currency", "BKPF-WAERS", 0, DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_CURRENCY, True)
Call VerifyTextBoxContent("/", "T001-WAERS", 0, DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_T001WAERS, True)
Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", 0, ConvertDate(DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCUMENT_DATE), True)
Call GetTextboxValue("BKPF-KURSF", 0, "DT_EXCHANGERATE_OUTPUT", True)
Call VerifyTextBoxContent("Exchange rate", "BKPF-KURSF", 0, TRIM(DT_EXCHANGERATE_OUTPUT), True)
Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)", True)
'
'''''''-----/ns_bce_68000174 screen--------------
'
Call SetTcode(DT_ZFIGL_UPLOAD_POST_750_OKCD)     
Call PressEnter()     
Call PressEnter()
Call SelectMenuBar("Selection;By Contents...")
Call SelectRowGuiTableByRow("SAPLSVIXTCTRL_SEL_FLDS", 1, True)
Call SelectRowGuiTableByRow("SAPLSVIXTCTRL_SEL_FLDS", 3, True)
Call SelectRowGuiTableByRow("SAPLSVIXTCTRL_SEL_FLDS", 4, True)
Call SelectRowGuiTableByRow("SAPLSVIXTCTRL_SEL_FLDS", 5, True)
Call TakeScreenShot
Call ClickButtonIfExist("Apply   \(Enter\))", True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call SetTableData("SAPLSVIXTCTRL_QUERY", "Field Contents", 1, "<NA>", "<NA>", DT_ZFIGL_UPLOAD_POST_600_TABLECELL_FIELD_CONTENTS_0, True)
Call SetTableData("SAPLSVIXTCTRL_QUERY", "Field Contents", 2, "<NA>", "<NA>", DT_ZFIGL_UPLOAD_POST_600_TABLECELL_FIELD_CONTENTS_1, True)
Call SetTableData("SAPLSVIXTCTRL_QUERY", "Field Contents", 3, "<NA>", "<NA>", DT_ZFIGL_UPLOAD_POST_600_TABLECELL_FIELD_CONTENTS_2, True)
Call SetTableData("SAPLSVIXTCTRL_QUERY", "Field Contents", 4, "<NA>", "<NA>", DT_ZFIGL_UPLOAD_POST_600_TABLECELL_FIELD_CONTENTS_3, True)
Call TakeScreenShot
Call ClickButtonIfExist("Confirm   \(Enter\)", True)
Call TakeScreenShot
Call VerifyTableCellContent(1, "Dir.quot.", "SAPL0SAPTCTRL_V_TCURR", DT_ZFIGL_UPLOAD_POST_20_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0)


'''''-----/nFB03 screen--------------

Call SetTcode(DT_ZFIGL_UPLOAD_POST_1400_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
Call VerifyTextBoxContent("Vendor", "RF05L-KTNRA", 0, DT_ZFIGL_UPLOAD_POST_303_CHECK_TEXT_OF_VENDOR, False)
Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_303_CHECK_TEXT_OF_GL_ACC, False)
Call ClickButtonIfExist("Display Next Item   \(Shift\+F7\)", False)
Call VerifyTextBoxContent("G/L Account", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_300_CHECK_TEXT_OF_GL_ACCOUNT, False)
Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, DT_ZFIGL_UPLOAD_POST_300_CHECK_TEXT_OF_ASSIGNMENT, False)
Call TakeScreenShot

Call LogOff()

Call FinalStatus ()

