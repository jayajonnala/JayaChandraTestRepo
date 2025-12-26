
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0014 Upload GL Document for reclassification PL account_Obsolete
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


gstrTestCaseName = "Test_GL0014 Upload GL Document for reclassification PL account_Obsolete"
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

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)   
Call SetTextbox("Session","P_SESS","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False) 
'Capture the screenshot
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False) 


Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
'Capture the screenshot
Call TakeScreenShot()
ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
Wait(2)

'Verify the details
Call VerifyGridCellContent("",1,"Amount",0,DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("",14,"Amount",0,DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_13_WRBTR)
Call VerifyGridCellContent("",1,"Document Header Text",0,DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BKTXT)
Call VerifyGridCellContent("",14,"Document Header Text",0,DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_13_BKTXT)

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
'----------------------Tcode SM35----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call TakeScreenShot()
Wait(2)

Call SetTextbox("Sess\.:","D0100-MAPN","",DT_ZFIGL_UPLOAD_POST_1005_SESS,False)    
Call SetTextbox("From:","D0100-VON","",ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM),False)    
Call SetTextbox("Created by:","D0100-CREATOR","",DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY,False)    
Call TakeScreenShot()
Wait(1)
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Call SelectColumnTable("SAPMSBDC_CCTC_APQI",0,"Session Name",False)
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
Call TakeScreenShot()

'Click execute
Call ClickButton("Process session   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
Call SelectRadioButtonByIndexIfPopupExists("D0300-ERROR",0)
Call TakeScreenShot()
Wait(2)
'Click execute
Call ClickButton("Process   \(Enter\)",True) 
Wait(5)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot()

'Click on Session overview"
Call ClickButton("Go back to batch input session overview   \(Enter\)",True) 
Wait(2)
Call TakeScreenShot()

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)


''Set obj = SAPGuiSession("activewindow:=Batch Input: Session Overview","program:=SAPMSBDC_CC").SAPGuiWindow("program:=SAPMSBDC_CC","text:=Batch Input: Session Overview")
''obj.SAPGuiTable("name:=SAPMSBDC_CCTC_APQI","index:=0").SelectRow 1
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)

'Set objWsh = CreateObject("WScript.Shell") 
'objWsh.SendKeys "{F2}"
'Set objWsh=nothing

' ClickButton(tooltipOrButtonName, blnIsItPopup)
Call ClickButton("Analyze session   \(F2\)",False)

'Verify the Status
Call GetTableCellData("RSBDC_ANALYSETC_TCODES","Status",1,"Trans","FB01","DT_STATUS_VAL",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))


'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DATE),False)
Wait(1)
Call TakeScreenShot()

'Verify the message
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",2,"Transaction","FB01","DT_MESSAGE_1",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(2,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_1))

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",4,"No.","363","DT_OUT_1",False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",5,"No.","364","DT_OUT_2",False)
''Call VerifyTableCellContent(4,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3))
''Call VerifyTableCellContent(5,"Message","RSBDC_ANALYSETC_PROTOCOL",LCase(DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4))
Call VerifyTableCellContent(4,"Message","RSBDC_ANALYSETC_PROTOCOL",DT_OUT_1)
Call VerifyTableCellContent(5,"Message","RSBDC_ANALYSETC_PROTOCOL",DT_OUT_2)

'----------------------Tcode FB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
Call TakeScreenShot()
Wait(2)

'Enter the details
Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_NUMBER,False)    
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_0100_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIGL_UPLOAD_POST_0100_FISCAL_YEAR,False)    
Call TakeScreenShot()
Wait(2)
Call PressEnter()     ' 

'Verify the content
Call VerifyGridCellContent("",1,"Amount",0,DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",14,"Amount",0,DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_13_AZBET)


'Click Document Display: General Ledger View
Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False) 
Wait(2)
Call TakeScreenShot()
Wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


'----------------------Tcode FAGLL03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC2) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)


Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_ZFIGL_UPLOAD_POST_1000_GL_ACCOUNT,False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False) 
Wait(2)
Call TakeScreenShot()
Wait(2)


'Click on GUI Tree
Call ActivateNodeGuiTree(0,"General Ledger Line Items;Document Number")
Wait(2)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_NUMBER_OCC1,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Yes",True)
wait(2)


'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

