'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0007 Manual posting of CN to be received (non Artemis) - 3rd p
'.................Author : TCS      
'................ Creation Date   
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

''earlier posting date was 10072014, changed to current date for execution

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


gstrTestCaseName = "Test_CL0007 Manual posting of CN to be received (non Artemis) - 3rd p"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0007 Manual posting of CN to be received (non Artemis) - 3rd p.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call RefreshExcelSheet(DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME)


''--------TransactionCode-ZFIGL_UPLOAD_POST---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("File path name","P_FILE", "", DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME, False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot()						
Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
If ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").exist(0) Then
	ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
''Capture the screenshot
Call TakeScreenShot()
ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
Wait(2)
End If

Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Reference", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR)
Call VerifyGridCellContent("", 2, "Reference", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XBLNR)
Call VerifyGridCellContent("", 3, "Reference", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XBLNR)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR)

Call VerifyGridCellContent("", 1, "NEWKO", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWKO)
Call VerifyGridCellContent("", 2, "NEWKO", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWKO)
Call VerifyGridCellContent("", 3, "NEWKO", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWKO)

Call VerifyGridCellContent("", 3, "MWSKZ", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("", 0, 3, False)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)
Call TakeScreenShot
''System/Services/Batch input/Sessions
Call SelectMenuBar("System;Services;Batch Input;Sessions")
Call TakeScreenShot

Call SetTextbox("Sess\.:","D0100-MAPN", "", DT_ZFIGL_UPLOAD_POST_1005_SESS, False)
Call SetTextbox("From:","D0100-VON", "", ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM), False)

Call SetTextbox("Created by:","D0100-CREATOR", "", DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY, False)
Call TakeScreenShot
Call PressEnter()     

Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Process session   \(F8\)", False)
Call TakeScreenShot

Call SelectRadioButtonIfExist("D0300-ERROR", 0, True)
Call TakeScreenShot
Call ClickButtonIfExist("Process   \(Enter\)", True)
Call TakeScreenShot
wait 1

Call ClickButtonIfExist("Go back to batch input session overview   \(Enter\)", True)
wait 1
Call TakeScreenShot

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call TakeScreenShot
Call ClickButtonIfExist("Log   \(F7\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Display   \(F2\)", False)
Call TakeScreenShot

wait 1
'''Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message", 3, "<NA>", "<NA>", "Message_OUTPUT", False)
Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message", 6, "<NA>", "<NA>", "Message_OUTPUT", False)

Call ClickButtonIfExist("Back   \(F3\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)", False)
Call TakeScreenShot


'''''''--------TransactionCode-/FB03----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)

Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_NUMBER,False)

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_0100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("", 3, "LOKKT", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LOKKT)

Call VerifyGridCellContent("", 1, "KOBEZ", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 2, "KOBEZ", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
Call VerifyGridCellContent("", 3, "KOBEZ", 0, lcase(DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOBEZ))

Call VerifyGridCellContent("", 1, "MWSKZ", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "MWSKZ", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 3, "MWSKZ", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)

Call VerifyGridCellContent("", 1, "AZBET", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 2, "AZBET", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 3, "AZBET", 0, trim(DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET))

Call VerifyGridCellContent("", 1, "RF05A_UBAZW", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)

Call DoubleClickGuiGridCell("", 0, 1, "RF05A_UBAZW", False)
Call TakeScreenShot

Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_0303_CHECK_TEXT_OF_GL_ACC, False)
Call ClickButtonIfExist("Call Up Document Overview   \(F9\)", False)
Call TakeScreenShot

Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 1, "Alternative Account No.", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_0303_CHECK_TEXT_OF_GL_ACC_OCC1, False)

Call ClickButtonIfExist("Call Up Document Overview   \(F9\)", False)
Call TakeScreenShot

Call SelectRowGuiGridbyRowNo("", 0, 3, False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 3, "Account", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, DT_ZFIGL_UPLOAD_POST_0300_CHECK_TEXT_OF_ASSIGNMENT, False)


Call LogOff()

Call FinalStatus ()

