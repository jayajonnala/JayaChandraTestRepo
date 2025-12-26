
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0013_Manage_Automatical_Customer_Invoicing_Crediting_Non-Po
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


gstrTestCaseName = "Test_AR0013_Manage_Automatical_Customer_Invoicing_Crediting_Non-Po"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode ZFIGL_UPLOAD_POST----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)   
Call SetTextbox("Session","P_SESS","","GTE089",False) 
'Capture the screenshot
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False) 


''Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
''ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
'''Capture the screenshot
''Call TakeScreenShot()
''ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
''Wait(2)


'Click execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)


'Click execute
Call ClickButton("Yes",True) 
Wait(2)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'----------------------Tcode SM35----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot()
Wait(2)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'Click execute
Call ClickButton("Process session   \(F8\)",False) 
Wait(2)
Call SelectRadioButtonByIndexIfPopupExists("D0300-BATCH",1)
Call TakeScreenShot()
Wait(2)
'Click execute
Call ClickButton("Process   \(Enter\)",True) 
Wait(2)

Call VerifyStatusBarMessageType("S")

'Verify the status bar message
Call VerifyStatusBar(DT_ZFIGL_UPLOAD_POST_1000_CHECK_TEXT_OF_STATUSBAR)
wait(15)
'Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'Wait(1)
'Set objWsh = CreateObject("WScript.Shell") 
'objWsh.SendKeys "{F2}" 
'Set objWsh=nothing
Call ClickButton("Log   \(F7\)",False)
Call ClickButton("Analyze session and logs   \(Shift\+F6\)",False) 

Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_DATE),False)
Wait(1)
Call TakeScreenShot()

'Verify the message

' FindRowNumber(tableNameOrGridTitle, refColumnName, refCellValue, dataTableColumnName)
Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL","No.",312,"ROW_NO_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("ROW_NO_OUTPUT",ROW_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",ROW_NO,"Transaction","FB01","DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4_OP",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4_OP",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


'----------------------Tcode FB03----------------------------
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call TakeScreenShot()
Wait(2)

'Enter the details
Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_100_DOCUMENT_NUMBER,False)    
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR_OCC1,False)    
Call TakeScreenShot()
Wait(2)
Call PressEnter()     

Call ClickContextButtonToolBar("&MB_VARIANT", "")
Call SelectMenuIdToolBar("&COL0", "")
Call ClickButtonToolBar("&FIND", "")
'''''Replaced above script instead of below one.
'''''
''''''SAPGuiSession("Session").SAPGuiWindow("Display Document: Data").SAPGuiToolbar("GridToolbar").PressContextButton "&MB_VARIANT"
''''''SAPGuiSession("Session").SAPGuiWindow("Display Document: Data").SAPGuiToolbar("GridToolbar").SelectMenuItemById "&COL0"
''''''SAPGuiSession("Session").SAPGuiWindow("Change Layout").SAPGuiToolbar("GridToolbar").PressButton "&FIND"

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZFIGL_UPLOAD_POST_100_GL_ACCOUNT,True)    
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)

' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
Call GetTextboxValue("BKPF-XBLNR",0,"DT_ZFIGL_UPLOAD_POST_750_CHECK_TEXT_OF_REFERENCE_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_750_CHECK_TEXT_OF_REFERENCE_OUTPUT",DT_ZFIGL_UPLOAD_POST_750_CHECK_TEXT_OF_REFERENCE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("",1,"G/L Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
'Commented below line as it has 2 G/L account columns in runtime applicaiton
'Call VerifyGridCellContent("",2,"G/L Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'Call VerifyGridCellContent("",1,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
'Call VerifyGridCellContent("",2,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"Assignment",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_TEXT_OF_REFERENCE)
Call VerifyGridCellContent("",2,"Assignment",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
Call VerifyGridCellContent("",1,"Tax Code",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("",2,"Tax Code",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)



'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

