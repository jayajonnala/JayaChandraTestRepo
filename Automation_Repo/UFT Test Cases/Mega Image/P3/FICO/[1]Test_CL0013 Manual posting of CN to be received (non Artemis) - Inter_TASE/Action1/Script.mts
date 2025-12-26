'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0013 Manual posting of CN to be received (non Artemis) - Inter  
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


gstrTestCaseName = "Test_CL0013 "
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0013 Manual posting of CN to be received (non Artemis) - Inter.xls"
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

'''Call SetTextbox("Sess\.:","D0100-MAPN", "", "", False)
'''Call TakeScreenShot
'''Call PressEnter()     

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
Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message", 3, "<NA>", "<NA>", "DT_ZFIGL_UPLOAD_POST_1400_GET_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT", False)
Call TakeScreenShot

'''''-----/nFB03 screen--------------

Call SetTcode(DT_ZFIGL_UPLOAD_POST_1400_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)


Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_1400_GET_TEXT_OF_TABLECELL_MESSAGE_3_OUTPUT,False)
Call SetTextbox("Company Code","RF05L-BUKRS","","RO02",False)
'Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call DoubleClickGuiGridCell("", 0, 1, "Account", False)
Call GetTextboxValue("BSEG-HKONT", 0, "DT_ZFIGL_UPLOAD_POST_303_GET_TEXT_OF_GL_ACC_OUTPUT_OUTPUT", False)
Call VerifyTextBoxContent("G/L Acc", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_303_GET_TEXT_OF_GL_ACC_OUTPUT_OUTPUT, False)
Call ClickButtonIfExist("Display Document Header   \(F5\)", False)
Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)", True)
Call ClickButtonIfExist("Display Next Item   \(Shift\+F7\)", False)
Call GetTextboxValue("BSEG-HKONT", 0, "DT_ZFIGL_UPLOAD_POST_300_CHECK_GET_OF_GL_ACCOUNT_OUTPUT_1", False)
Call VerifyTextBoxContent("G/L Account", "BSEG-HKONT", 0, DT_ZFIGL_UPLOAD_POST_300_CHECK_GET_OF_GL_ACCOUNT_OUTPUT_1, False)
Call VerifyTextBoxContent("Assignment", "BSEG-ZUONR", 0, DT_ZFIGL_UPLOAD_POST_300_GET_TEXT_OF_ASSIGNMENT, False)
Call TakeScreenShot

Call LogOff()

Call FinalStatus ()

