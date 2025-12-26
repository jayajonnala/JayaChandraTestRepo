
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AR0012_Manage_Automatical_Customer_Invoicing_Crediting_Non_Po_In
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


gstrTestCaseName = "Test_AR0012_Manage_Automatical_Customer_Invoicing_Crediting_Non_Po_In"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''----------------------Tcode ZFIGL_UPLOAD_POST----------------------------
'
''Enter the Tcode
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

'Click execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)


'Click execute
Call ClickButton("Yes",True) 
Wait(2)

'Capture the screenshot
Call TakeScreenShot()

Call  VerifyifGuiLabelExists(Lcase(DT_ZFIGL_UPLOAD_CHECK_POST_GTE089_WAS_OPENED))
Call  VerifyifGuiLabelExists(Lcase(DT_ZFIGL_UPLOAD_POST_CHECK_GTE089_WAS_CREATED))

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call CheckTCodeScreen("SESSION_MANAGER")

''----------------------Tcode SM35----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ZFIGL_UPLOAD_POST_120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
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
'
''Verify the status bar message
Call VerifyStatusBar(DT_ZFIGL_UPLOAD_POST_1000_CHECK_TEXT_OF_STATUSBAR)
'
Call ClickButton("Log   \(F7\)",False)
Call ClickButton("Analyze session and logs   \(Shift\+F6\)",False) 
'Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
'Wait(1)
'Set objWsh = CreateObject("WScript.Shell") 
'objWsh.SendKeys "{F2}" 
'Set objWsh=nothing
'
'
Call VerifyTableCellContent(1,"Status","RSBDC_ANALYSETC_TCODES",LCase("Processed"))
''
''
''Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_DATE),False)
Wait(1)
Call TakeScreenShot()

'Verify the message
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",3,"Transaction","FB01","DT_OUTPUT_MESSAGE_1",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


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
Call SetTextbox("Document Number","RF05L-BELNR","",DT_DOC_NO,False)    
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_COMP_CODE,False)    
'Call SetTextbox("Fiscal Year","RF05L-GJAHR","","2021",False)    
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FISCAL_YEAR,False)    
Call TakeScreenShot()
Wait(2)
Call PressEnter()     ' 

Call GetTextboxValue("BKPF-XBLNR",0,"DT_REFEReNCE_NO_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Verify the Posting Key Values
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"Posting Key",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("",1,"Account",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("",1,"Assignment",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"Assignment",0,DT_FB03_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)


'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************
