'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Upload AR Accounting Document_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 18th Feb
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Upload AR Accounting Document_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Upload AR Accounting Document_p1.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call SetTcode(DT_T_CODE) 
Call PressEnter()  
'Call SelectMenuBar("Settings;Processing Options")
Call SelectMenuBar("Settings;Editing Options")
Call TakeScreenShot()
Call SelectCheckbox("RFOPT2-XSNET", 0, "OFF", False)
Call TakeScreenShot()
Call ClickButton("Change user master   \(Ctrl\+S\)", False)
Call TakeScreenShot()
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

'Capture the screenshot
Call TakeScreenShot()
'Click execute
Call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyGridCellContent("",1,"Document Number","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZZZ_KEY)
Call VerifyGridCellContent("",1,"Company Code","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
Call VerifyGridCellContent("",1,"Date","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT)
Call VerifyGridCellContent("",1,"Posting Key","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWBS)
Call VerifyGridCellContent("",1,"Amount","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)

' SelectRowGuiGridbyRowNo(gridTitle, gridIndex, rowno, blnIsItPopup)
Call SelectRowGuiGridbyRowNo("","",1,False)
'set selected rows
'call SelectRowGuiGrid("","","Document Number",1,False)
'Capture the screenshot
Call TakeScreenShot()
call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()
call ClickButton("Yes",True)
wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call SetTcode(DT_T_CODE) 
Call PressEnter()  
'Call SelectMenuBar("Settings;Processing Options")
Call SelectMenuBar("Settings;Editing Options")
Call TakeScreenShot()
Call SelectCheckbox("RFOPT2-XSNET", 0, "ON", False)
Call TakeScreenShot()
Call ClickButton("Change user master   \(Ctrl\+S\)", False)
Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

