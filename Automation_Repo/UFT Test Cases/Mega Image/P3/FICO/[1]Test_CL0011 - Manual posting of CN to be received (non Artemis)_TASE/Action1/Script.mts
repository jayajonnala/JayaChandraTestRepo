'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0011 - Manual posting of CN to be received (non Artemis)
'.................Author : TCS      
'................ Creation Date   
'.................Modified By :
'.................Modified Date/Details :
'
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


gstrTestCaseName = "Test_CL0011 - Manual posting of CN to be received (non Artemis)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0011 - Manual posting of CN to be received (non Artemis).xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-OB08---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call ClickButtonIfExist("VIM_POSI_PUSH", False)
'''Call SetTextbox("Exch.rate type","SVALD-VALUE", "", DT_OB08_0300_EXCH_RATE_TYPE, True)
Call SetTextboxNoLabel("SVALD-VALUE", "", DT_OB08_0300_EXCH_RATE_TYPE, True)
Call SetTextbox("From currency","SVALD-VALUE", "", DT_OB08_0300_FROM_CURRENCY, True)
Call SetTextbox("To-currency","SVALD-VALUE", "", DT_OB08_0300_TOCURRENCY, True)
Call SetTextbox("Valid from","SVALD-VALUE", "", ConvertDate(DT_OB08_0300_VALID_FROM), True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call GetTableCellData("SAPL0SAPTCTRL_V_TCURR", "Dir.quot.", 1, "<NA>", "<NA>", "DT_OB08_0020_CHECK_TEXT_OF_TABLECELL_DIRQUOT_0_OUTPUT", False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)

''''--------TransactionCode-ZFIGL_UPLOAD_POST---------''''

Call SetTcode(DT_OB08_0100_OKCD)     
Call PressEnter()  
Call SetTextbox("Session","P_SESS", "", DT_OB08_1000_SESSION, False)
Call SetTextbox("File path name","P_FILE", "", DT_OB08_1000_FILE_PATH_NAME, False)
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


Call VerifyGridCellContent("", 1, "WRBTR", 0, DT_OB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("", 2, "WRBTR", 0, DT_OB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)
Call VerifyGridCellContent("", 1, "NEWKO", 0, DT_OB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWKO)
Call VerifyGridCellContent("", 2, "NEWKO", 0, DT_OB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWKO)
Call TakeScreenShot()						
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call ClickButtonIfExist("Yes", True)
Call PressEnter()
Call ClickButtonIfExist("Back   \(F3\)", False)
Call SelectMenuBar("System;Services;Batch Input;Sessions")
Call SetTextbox("Sess.:","D0100-MAPN", "", DT_OB08_1005_SESS, False)
Call SetTextbox("From:","D0100-VON", "", ConvertDate(DT_OB08_1005_FROM), False)
Call SetTextbox("Created by:","D0100-CREATOR", "", DT_OB08_1005_CREATED_BY, False)
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButtonIfExist("Process session   \(F8\)", False)
Call SelectRadioButtonIfExist("D0300-ERROR", 0, True)
Call ClickButtonIfExist("Process   \(Enter\)", True)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Sess.:","D0100-MAPN", "", DT_OB08_1005_SESS, False)
Call SetTextbox("From:","D0100-VON", "", ConvertDate(DT_OB08_1005_FROM), False)
Call SetTextbox("Created by:","D0100-CREATOR", "", DT_OB08_1005_CREATED_BY, False)
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Log   \(F7\)", False)
Call ClickButtonIfExist("Display   \(F2\)", False)
wait 1

Call FindRowNumber("RSBDC_PROTOCOLTC_PROT_DIS", "Message Number", "312", "DT_RowNumber")

Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message", DT_RowNumber, "", "", "DT_OB08_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2_OUTPUT", False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)

''''''-----/nFB03 screen--------------

Call SetTcode(DT_OB08_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_OB08_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_OB08_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_OB08_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
Call VerifyTextBoxContent("Amount", "BSEG-WRBTR", 0, DT_OB08_0303_CHECK_TEXT_OF_AMOUNT, False)
Call VerifyTextBoxContent("LC amount", "BSEG-DMBTR", 0, DT_OB08_0303_CHECK_TEXT_OF_LC_AMOUNT, False)
Call VerifyTextBoxContent("Text", "BSEG-SGTXT", 0, DT_OB08_0303_CHECK_TEXT_OF_TEXT, False)
Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_OB08_0303_CHECK_TEXT_OF_GL_ACC, False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)

''''''-----//nFBL3N screen--------------

Call SetTcode(DT_OB08_0100_OKCD_OCC2)     
Call PressEnter()   

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_OB08_1000_GL_ACCOUNT,False)
Call SelectRadioButton("X_AISEL", "All items", False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_OB08_1000_COMPANY_CODE,False)

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)", False)
Call ActivateNodeGuiTree(0, "#3;#4")
Call ClickButtonToolBar("TAKE", 0)
Call ClickButtonIfExist("%_%%DYN007_%_APP_%-VALU_PUSH", False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_OB08_3010_TABLECELL_SINGLE_VALUE_0, True)
Call ClickButtonIfExist("Copy   \(F8\)", True)
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call VerifyifGuiLabelExists(DT_OB08_0120_CHECK_TEXT_OF_KA)
Call TakeScreenShot
Call ClickButtonIfExist("Display Document   \(Shift\+F2\)", False)
Call TakeScreenShot
'Call VerifyTextBoxContent("Exchange rate", "BKPF-KURSF", 0, trim(DT_OB08_1710_CHECK_TEXT_OF_EXCHANGE_RATE), True)
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)", True)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)

Call LogOff()

Call FinalStatus ()

