
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


gstrTestCaseName = "Test_06-01-02-13-04-Create CPO for LUX store_p2"
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
'--------------------------------------------  ZMDPU_COLLECT_PO----------------------------------------------


Call SetTextbox("Table Name","DATABROWSE-TABLENAME","",DT_SE16_0230_TABLE_NAME,False)
Call PressEnter()
'Call SetTextbox("Original Doc\.Number","I3-LOW","",DT_SE16_1000_ORIGINAL_DOCNUMBER,False)
Call SetTextbox("DOCNR","I3-LOW","",DT_SE16_1000_ORIGINAL_DOCNUMBER,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",false)
Call TakeScreenShot()

'Call GetGridContent("",0, "EBELN", 1, "<NA>", "<NA>", "DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OUTPUT")
' GetLabelContentByRefLabel(refLabelContent, xDifferenceValue, yDifferenceValue, dataTableColumnName, blnIsItPopup)
Call GetLabelContentByRefLabel("EBELN",0,-32, "DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OUTPUT",false)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyGridCellContent("",1,"EBELN",1,DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OCC1)
' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
'Call VerifyifGuiLabelExistsByRelativeid(DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OCC1,wnd\[0\]/usr/lbl\[85,5\])



'Call GetGridContent("", 0, "EBELN", 1, "<NA>", "<NA>", "DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OUTPUT")
' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)
'call GetTableCellData("", "EBELN",1,"<NA>", "<NA>","DT_EBELN",false)
' GetLabelContentByRefLabel(refLabelContent, xDifferenceValue, yDifferenceValue, dataTableColumnName, blnIsItPopup)
'call GetLabelContentByRefLabel(Replace(date,"/","."),259,0,"DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OUTPUT",false)
'call GetLabelContentByRefLabel(Replace(date,"/","."),0,-32,"DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OUTPUT",false)
' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)


Call LogOff()
Call FinalStatus ()



