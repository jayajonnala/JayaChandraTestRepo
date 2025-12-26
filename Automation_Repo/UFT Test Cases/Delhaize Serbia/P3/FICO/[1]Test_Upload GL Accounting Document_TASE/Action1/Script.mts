'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Upload GL Accounting Document_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 3rd May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Upload GL Accounting Document_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Upload GL Accounting Document_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode ZFIGL_UPLOAD_POST----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)
Call SetTextbox("Session","P_SESS","",DT_ZFIGL_UPLOAD_POST_1000_SESSION,False)
'Capture the screenshot
Call TakeScreenShot()
'Click execute
Call ClickButton("Execute   \(F8\)",False)
Wait(5)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Posting Key","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWBS)
Call VerifyGridCellContent("",2,"Posting Key","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWBS)
Call VerifyGridCellContent("",4,"Posting Key","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NEWBS)

Call VerifyGridCellContent("",1,"Account","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWKO)
Call VerifyGridCellContent("",2,"Account","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWKO)
Call VerifyGridCellContent("",3,"Account","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWKO)
Call VerifyGridCellContent("",4,"Account","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NEWKO)

Call VerifyGridCellContent("",1,"Cost Center","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",2,"Cost Center","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("",3,"Cost Center","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
Call VerifyGridCellContent("",4,"Cost Center","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PRCTR)

Call VerifyGridCellContent("",1,"Amount","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("",2,"Amount","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)
Call VerifyGridCellContent("",3,"Amount","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR)
Call VerifyGridCellContent("",4,"Amount","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_WRBTR)

'set selected rows
call SelectRowGuiGrid("","","Document Number",1,False)
'Capture the screenshot
Call TakeScreenShot()
call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()
call ClickButtonIfExist("Yes",True)
wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call SelectMenuBar("System;Services;Batch Input;Sessions")
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'select row
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",DT_ZFIGL_UPLOAD_POST_0500_GRIDCELL_0_DOCUMENTNO,False)
'Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name","GTE089",False)
Call ClickButton("Process session   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call SelectRadioButtonIfPopupExists("D0300-ERROR",DT_ZFIGL_UPLOAD_POST_0300_DISPLAY_ERRORS_ONLY)
Call ClickButton("Process   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()

For Iterator = 1 To 2 Step 1
	Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_POSTING_DATE,"/","."),False)
    Call PressEnter()
    Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_DOCUMENT_DATE,"/","."),False)
    Call PressEnter()
Next
'Capture the screenshot
call TakeScreenShot()

Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",DT_ZFIGL_UPLOAD_POST_0500_GRIDCELL_0_DOCUMENTNO,False)
'Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name","GTE089",False)
'Capture the screenshot
call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


