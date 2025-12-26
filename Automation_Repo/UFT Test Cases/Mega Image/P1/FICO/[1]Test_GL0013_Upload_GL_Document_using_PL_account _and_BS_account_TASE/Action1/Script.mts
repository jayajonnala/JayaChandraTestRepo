
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0013_Upload_GL_Document_using_PL_account _and_BS_account_TASE
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


gstrTestCaseName = "Test_GL0013_Upload_GL_Document_using_PL_account _and_BS_account_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''----------------------Tcode ZFIGL_UPLOAD_POST----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)   
Call SetTextbox("Session","P_SESS","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False) 
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False) 

Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
Call TakeScreenShot()
ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
Wait(2)
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)

Call ClickButton("Yes",True) 
Wait(2)

Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
wait(2)
Call ClickButton("Back   \(F3\)",False)
wait(2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'----------------------Tcode SM35----------------------------

Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot()
Wait(2)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'Click execute
Call ClickButton("Process session   \(F8\)",False) 
Wait(2)
Call SelectRadioButtonByIndexIfPopupExists("D0300-BATCH",0)
Call TakeScreenShot()
Wait(2)
'Click execute
Call ClickButton("Process   \(Enter\)",True) 
Wait(2)

'Verify the status bar message
Call VerifyStatusBar(Lcase(DT_ZFIGL_UPLOAD_POST_1000_CHECK_TEXT_OF_STATUSBAR))


'Click on Log
Call ClickButton("Log   \(F7\)",False) 
Wait(2)

'Click on Analyze session and logs
Call ClickButton("Analyze session and logs   \(Shift\+F6\)",False) 
Wait(2)

'Verify the Status
Call GetTableCellData("RSBDC_ANALYSETC_TCODES","Status",1,"Trans","FB01","DT_TMP_VAL",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))


'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DATE),False)
Wait(1)
Call TakeScreenShot()

'Verify the message
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",3,"Transaction","FB01","DT_MESSAGE_1",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(3,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_8))

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


'----------------------Tcode FB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ZFIGL_UPLOAD_POST_100_OKCD_OCC2)
Call TakeScreenShot()
Wait(2)

'Enter the details
Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_100_DOCUMENT_NUMBER,False)    
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_100_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_100_FISCAL_YEAR,False)    
Call TakeScreenShot()
Wait(2)
Call PressEnter()     ' 


'Click Display Document Header   \(F5\)
Call ClickButton("Display Document Header   \(F5\)",False) 
Wait(2)

'Enter the details
Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,True)
Call VerifyTextBoxContent("Doc\.Header Text","BKPF-BKTXT",0,LCase(DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCHEADER_TEXT),True)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR",0,DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_SESSION_NAME,True)
Call TakeScreenShot()
Wait(2)


'Click Continue/Confirm   \(Enter\)
Call ClickButton("Continue/Confirm   \(Enter\)",True) 
Wait(2)

'Select the G/L Account column and press Find
Call SelectColumnGuiGrid("",0,"G/L Account",False)
Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","G/L Account",True) 
Call TakeScreenShot()
Wait(2)
Call ClickButtonIfExist("OK   \(Enter\)",True)
Call ClickButtonIfExist("Cancel   \(F12\)",True)


'Click Document Display: General Ledger View
Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False) 
Wait(2)
Call TakeScreenShot()
Wait(2)

Call VerifyGridCellContent("",1,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("",2,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("",3,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyGridCellContent("",4,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)
Call VerifyGridCellContent("",5,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_HKONT)
Call VerifyGridCellContent("",6,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_HKONT)
Call VerifyGridCellContent("",7,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_HKONT)
Call VerifyGridCellContent("",8,"Account",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_HKONT)

'Verify the Posting Key Values
Call VerifyGridCellContent("",1,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("",4,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("",5,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("",6,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("",7,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BSCHL)
Call VerifyGridCellContent("",8,"Posting Key",0,DT_ZFIGL_UPLOAD_POST_750_CHECK_GETCELLVALUE_OF_GRIDCELL_7_BSCHL)

'Click Select Other Ledger
Call ClickButton("Select Other Ledger   \(Ctrl\+F8\)",False) 
Wait(2)
Call TakeScreenShot()
Wait(2)

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************
