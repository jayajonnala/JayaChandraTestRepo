

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Direct Export to foreign customers_p3
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Direct Export to foreign customers_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_AT_Direct Export to foreign customers_p3_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
'----------------------VF01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 
Call TakeScreenShot

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT", "Document", 1, "", "", DT_VF01_0102_TABLECELL_DOCUMENT_0, false)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1", "DT_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_STATUSBAR_OUTPUT&" has been saved")


'-----------------------------VF03-----------------------------
Call SetTcode(DT_VF01_0102_OKCD)
Call PressEnter() 
Call TakeScreenShot

Call SelectMenuBar("Billing document;Display")
Call TakeScreenShot
Call SetTextbox("Billing document","VBRK-VBELN","",DT_STATUSBAR_OUTPUT,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Accounting overview   \(Shift\+F4\)",False)
Call TakeScreenShot
'verify accounting document grid
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Document Number", True)
Call TakeScreenShot
Call VerifyGridCellContent("", 2, "Profit Center", "", DT_VF01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
Call GetTextboxValue("BKPF-BELNR","","DT_VF01_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT",False)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call SelectMenuBar("Billing document;Issue Output To")
Call TakeScreenShot
Call VerifyTableCellContent(1,"Message Type","SAPLVMSGTABCONTROL",DT_VF01_0200_CHECK_TEXT_OF_TABLECELL_MESSAGE_TYPE_0)

Call SelectRowGuiTable("SAPLVMSGTABCONTROL","Message Type",DT_VF01_0200_CHECK_TEXT_OF_TABLECELL_MESSAGE_TYPE_0,True)
Call ClickButton("Print preview   \(Ctrl\+Shift\+F1\)",True)
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
