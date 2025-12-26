'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0001 Manual posting of CN to be received (non Artemis) _NEW 
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


gstrTestCaseName = "Test_CL0001 Manual posting of CN to be received (non Artemis) _NEW"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0001 Manual posting of CN to be received (non Artemis)_NEW.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-ZFIGL_UPLOAD_POST---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Session","P_SESS", "", DT_ZFIGL_UPLOAD_POST_1000_SESSION, False)
Call SetTextbox("File path name","P_FILE", "", DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME, False)
'''wait 2
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)
wait 2
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

''-----SM35 screen--------------
Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

''DT_ZFIGL_UPLOAD_POST_1005_SESS
Call SetTextbox("Sess\.:","D0100-MAPN", "", "", False)
Call TakeScreenShot
Call PressEnter()     

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)

Call ClickButtonIfExist("Process session   \(F8\)", False)

Call SelectRadioButtonIfExist("D0300-ERROR", 0, True)
Call TakeScreenShot
Call ClickButtonIfExist("Process   \(Enter\)", True)
'Call TakeScreenShot
Call PressEnter()
Call PressEnter()


Call ClickButtonIfExist("Go back to batch input session overview   \(Enter\)", True)
wait 1
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)

Call ClickButtonIfExist("Log   \(F7\)", False)
Call ClickButtonIfExist("Display   \(F2\)", False)
wait 1
Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message", 1, "<NA>", "<NA>", "Message_OUTPUT", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)
Call ClickButtonIfExist("Back   \(F3\)", False)

	
''''-----FB03 screen--------------
Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)
	

'Call SetTextbox("Company Code","RF05L-BUKRS", "", DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE, False)
'Call SetTextbox("Fiscal Year","RF05L-GJAHR", "", DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR, False)
Call PressEnter()
Call TakeScreenShot
'' Verify gridcell content
'' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("", 1, "Account", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "Account", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

'Assignment
Call VerifyGridCellContent("", 1, "Assignment", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 2, "Assignment", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
Call VerifyGridCellContent("", 3, "Assignment", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)
Call VerifyGridCellContent("", 4, "Assignment", 0, DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ZUONR)

Call ClickButtonIfExist("Display Document Header   \(F5\)", False)
Call ClickButtonIfExist("Continue/Confirm   \(Enter\)", True)

''Doubleclick current cell
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)

Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_303_CHECK_TEXT_OF_GL_ACC, False)
Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, DT_ZFIGL_UPLOAD_POST_303_CHECK_TEXT_OF_ASSIGNMENT, False)

Call ClickButtonIfExist("Display Next Item   \(Shift\+F7\)", False)

Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_303_CHECK_TEXT_OF_GL_ACC_OCC2, False)
Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, DT_ZFIGL_UPLOAD_POST_303_CHECK_TEXT_OF_ASSIGNMENT_OCC2, False)
Call TakeScreenShot
Call ClickButtonIfExist("Display Next Item   \(Shift\+F7\)", False)

Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, DT_ZFIGL_UPLOAD_POST_300_CHECK_TEXT_OF_ASSIGNMENT, False)

Call ClickButtonIfExist("Display Next Item   \(Shift\+F7\)", False)

Call VerifyTextBoxContent("G/L Account", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_300_CHECK_TEXT_OF_GL_ACCOUNT, False)
Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, lcase(DT_ZFIGL_UPLOAD_POST_300_CHECK_TEXT_OF_ASSIGNMENT_OCC2), False)

Call TakeScreenShot


Call LogOff()

Call FinalStatus ()

