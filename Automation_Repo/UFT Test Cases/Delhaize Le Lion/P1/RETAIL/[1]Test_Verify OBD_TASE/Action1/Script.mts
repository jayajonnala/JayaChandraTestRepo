
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.46 VIM - PO Precontrole Issue - BR24 - Missing Posting Date (PO)
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_Verify OBD"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'all LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
''
'--------------------------------------------  VL03N----------------------------------------------
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()

' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Ship-To Party","KUWEV-KUNNR","",DT_VL03N_1502_CHECK_TEXT_OF_SHIPTO_PARTY,False)
' VerifyTableCellContent(tableRowNumber, tableColumnName, tableName, expectedValue)
call VerifyTableCellContent(1,"Article","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
call VerifyTableCellContent(2,"Article","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)

call VerifyTableCellContent(1,"Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_0)
call VerifyTableCellContent(2,"Deliv. Qty","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_1)
'call VerifyTableCellContent(1,"Quantity Actually Delivered","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_0)
'call VerifyTableCellContent(2,"Quantity Actually Delivered","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_DELIV_QTY_1)

call VerifyTableCellContent(1,"Site","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SITE_0)
call VerifyTableCellContent(2,"Site","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SITE_1)

'call VerifyTableCellContent(1,"S.Loc","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_0)
'call VerifyTableCellContent(2,"S.Loc","SAPMV50ATC_LIPS_OVER",DT_VL03N_1102_CHECK_TEXT_OF_TABLECELL_SLOC_1)


Call LogOff()
Call FinalStatus ()



