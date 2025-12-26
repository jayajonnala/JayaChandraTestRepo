'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0002 Manual posting of CN to be received- warning for missing RVA code 
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


gstrTestCaseName = "Test_CL0002 Manual posting of CN to be received- warning for missing RVA code"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0002.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''--------TransactionCode-ZFIGL_UPLOAD_POST---------''''
'
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



Call VerifyGridCellContent("", 1, "Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWKO)
Call VerifyGridCellContent("", 2, "Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWKO)
Call VerifyGridCellContent("", 3, "Account", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWKO)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("", 0, 3, False)
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call ClickButtonIfExist("Yes", True)

Call SelectMenuBar("System;Services;Batch Input;Sessions")
wait 1
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
Call SetTextbox("Sess\.:","D0100-MAPN", "", DT_ZFIGL_UPLOAD_POST_1005_SESS, False)
Call SetTextbox("From:","D0100-VON", "", ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM), False)
Call SetTextbox("To:","D0100-BIS", "", ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_TO_), False)
	
Call SetTextbox("Created by:","D0100-CREATOR", "", DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY, False)
     

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)

Call ClickButtonIfExist("Process session   \(F8\)", False)

Call SelectRadioButtonIfExist("D0300-ERROR", 0, True)
Call TakeScreenShot
Call ClickButtonIfExist("Process   \(Enter\)", True)
Call TakeScreenShot

wait 1
Call VerifyTextBoxContent("Information Message", "MESSTXT1", 0, lcase("Processing of batch input session completed"), True)

Call LogOff()

Call FinalStatus ()

