
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0014_Manage_Automatical_Customer_Invoicing_Crediting_Non_Po_In
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


gstrTestCaseName = "Test_AR0014_Manage_Automatical_Customer_Invoicing_Crediting_Non_Po_In"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'----------------------Tcode ZFIGL_UPLOAD_POST----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME1,False)   
Call SetTextbox("Session","P_SESS","","GTE089",False) 
'Capture the screenshot
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False) 


'Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
'ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
''Capture the screenshot
'Call TakeScreenShot()
'ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
'Wait(2)


'Click execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)


'Click execute
Call ClickButtonIfExist("Yes",True) 
Wait(2)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)



'----------------------Tcode SM35----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_TRANSACTION) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_TRANSACTION)
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

Call ClickButton("Log   \(F7\)",False)
Call ClickButton("Analyze session and logs   \(Shift\+F6\)",False) 
'
'Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'Wait(1)
'Set objWsh = CreateObject("WScript.Shell") 
'objWsh.SendKeys "{F2}" 
'Set objWsh=nothing
'Wait(5)
'
Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_DATE),False)
Wait(1)
Call TakeScreenShot()

' FindRowNumber(tableNameOrGridTitle, refColumnName, refCellValue, dataTableColumnName)
Call FindRowNumber("RSBDC_ANALYSETC_PROTOCOL","No.",312,"ROW_NO_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("ROW_NO_OUTPUT",ROW_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''Verify the message
''Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",2,"Transaction","FB01","DT_MESSAGE_1",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",ROW_NO,"Transaction","FB01","DT_MESSAGE_1",False)

'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


'----------------------Tcode FB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot()
Wait(2)

'Enter the details
Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_100_DOCUMENT_NUMBER,False)    
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR,False)    
Call TakeScreenShot()
Wait(2)
Call PressEnter() 


Call VerifyGridCellContent("",3,"Negative Posting",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XNEGP)
Call VerifyGridCellContent("",1,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("",1,"Assignment",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"Assignment",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
Call VerifyGridCellContent("",1,"Text",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("",2,"Text",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("",1,"Tax Code",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("",2,"Tax Code",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("",3,"Tax Code",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)

'Click Display Document Header   \(F5\)
Call ClickButton("Display Document Header   \(F5\)",False) 
Wait(2)

'Enter the details
Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
Call VerifyTextBoxContent("Doc\.Header Text","BKPF-BKTXT",0,LCase(DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCHEADER_TEXT),True)
Call GetTextboxValue("BKPF-XBLNR",0,"DT_REFERENCE_OUTPUT",True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyTextBoxContent("Reference","BKPF-XBLNR",0,DT_REFERENCE,True)
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT",0,ConvertDate(DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCUMENT_DATE),True)
Call VerifyTextBoxContent("Reference Key","BKPF-AWKEY",0,DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_REFERENCE_KEY,True)
Call TakeScreenShot()
Wait(2)

'Click Continue/Confirm   \(Enter\)
Call ClickButton("Continue/Confirm   \(Enter\)",True) 
Wait(2)


'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************
