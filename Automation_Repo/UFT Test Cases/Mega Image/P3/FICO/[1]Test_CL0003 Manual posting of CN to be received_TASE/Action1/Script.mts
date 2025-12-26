'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_CL0003 Manual posting of CN to be received 
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


gstrTestCaseName = "Test_CL0003 Manual posting of CN to be received"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\FICO\TASE_DT_CL0003 Manual posting of CN to be received.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''----------------------Login----------------------------

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''--------TransactionCode-ZFIGL_UPLOAD_POST---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)



Call SetTextbox("File path name","P_FILE", "", DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME, False)
''Call SetTextbox("Session","P_SESS", "", DT_ZFIGL_UPLOAD_POST_1000_SESSION, False)
''wait 2
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

'' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)

Call VerifyGridCellContent("", 1, "Document type", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 2, "Document type", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 3, "Document type", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART_OCC1)

Call VerifyGridCellContent("", 1, "Reference", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR)
Call VerifyGridCellContent("", 2, "Reference", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XBLNR)
Call VerifyGridCellContent("", 3, "Reference", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XBLNR)

Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWBS)
Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWBS)
Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWBS)

Call VerifyGridCellContent("", 1, "Amount", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_WRBTR)

Call VerifyGridCellContent("", 3, "Assignment", 0, DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)
Call TakeScreenShot


Call SelectRowGuiGridbyRowNo("", 0, DT_ZFIGL_UPLOAD_POST_0500_GRIDCELL_2_DOCUMENTNO_OCC1, False)

Call ClickButtonIfExist("Execute   \(F8\)", False)
Call ClickButtonIfExist("Yes", True)
''System/Services/Batch input/Sessions
Call SelectMenuBar("System;Services;Batch Input;Sessions")
wait 1
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Sess\.:","D0100-MAPN", "", DT_ZFIGL_UPLOAD_POST_1005_SESS, False)
Call SetTextbox("From:","D0100-VON", "", ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM), False)
Call SetTextbox("Created by:","D0100-CREATOR", "", DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY, False)
Call TakeScreenShot
Call PressEnter()     

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)

Call ClickButtonIfExist("Process session   \(F8\)", False)


Call SelectRadioButtonIfExist("D0300-ERROR", 0, True)
Call TakeScreenShot
Call ClickButtonIfExist("Process   \(Enter\)", True)
Call TakeScreenShot

Call VerifyTextBoxContent("Information Message", "MESSTXT1", 0, lcase("Processing of batch input session completed"), True)


Call LogOff()

Call FinalStatus ()

