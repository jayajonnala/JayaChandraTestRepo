
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06-09-04-transfer (rebaptême) to bloc (WMS code R)_P2_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06-09-04-transfer (rebaptême) P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_04-01-01-01-01-05-Send PO to WMS_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''--------LOGIN----------''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''--------TransactionCode-MB51----------''''
 
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE) 


' SetTextbox(textboxAttachedText, textboxName, textboxIndex, textboxValue, blnIsItPopup)
Call SetTextbox("Article","MATNR-LOW","",DT_MB51_1000_ARTICLE,False)
Call SetTextbox("Site","WERKS-LOW","",DT_MB51_1000_SITE,False)
'Call SetTextbox("Company Code","BUKRS-LOW","",DT_MB51_1000_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BUDAT-LOW","",ConvertDate(DT_MB51_1000_POSTING_DATE),False)
Call SetTextbox("User Name","USNAM-LOW","","",False)
Call SetTextbox("Purchase order","EBELN-LOW","","",False)
Call TakeScreenShot() 

' FocusTextBox(attachedText, textboxName, blnIsItPopup)
'Call FocusTextBox("Purchase order","EBELN-LOW",False)

' ClickButton(tooltipOrButtonName, blnIsItPopup)
'Call ClickButton("%_EBELN_%_APP_%-VALU_PUSH",False)

' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
'Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_MB51_3010_TABLECELL_SINGLE_VALUE_0,True)
'Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_MB51_3010_TABLECELL_SINGLE_VALUE_1,True)
'Call TakeScreenShot() 
'Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)

Call ClickButton("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)

SAPGuiSession("Session_2").SAPGuiWindow("Change Layout").SAPGuiToolbar("GridToolbar").PressButton "&FIND"

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM,True)
Call TakeScreenShot() 
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot() 

Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)

SAPGuiSession("Session_2").SAPGuiWindow("Change Layout").SAPGuiToolbar("GridToolbar").PressButton "&FIND"

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot() 
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot() 

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)

SAPGuiSession("Session_2").SAPGuiWindow("Define Filter Criteria").SAPGuiToolbar("GridToolbar").PressButton "&FIND"

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_MB51_0841_SEARCH_TERM,True)
Call TakeScreenShot() 
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)

Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call ClickButton("Multiple selection",True)

Call SetTableData("SAPLALDBSINGLE","Single value",1,"","",DT_MB51_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","",DT_MB51_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot() 
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot() 

' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
Call GetGridContent("",0,"Article Document",1,"Document Header Text",DT_MB51_3010_TABLECELL_SINGLE_VALUE_1,"DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OUTPUT")
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Posting Date",0,ConvertDate(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))
Call VerifyGridCellContent("",1,"Movement type",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
Call VerifyGridCellContent("",1,"Article",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",1,"Document Header Text",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BKTXT)
Call VerifyGridCellContent("",1,"Reason for Movement",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GRUND)

Call GetGridContent("",0,"Article Document",3,"Document Header Text",DT_MB51_3010_TABLECELL_SINGLE_VALUE_0,"DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MBLNR_OUTPUT")

Call VerifyGridCellContent("",2,"Posting Date",0,ConvertDate(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUDAT))
Call VerifyGridCellContent("",2,"Document Header Text",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BKTXT)
Call VerifyGridCellContent("",2,"Reason for Movement",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GRUND)

Call GetGridContent("",0,"Article Document",5,"Document Header Text",DT_MB51_3010_TABLECELL_SINGLE_VALUE_0,"DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MBLNR_OUTPUT")

Call VerifyGridCellContent("",5,"Posting Date",0,ConvertDate(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BUDAT))
Call VerifyGridCellContent("",5,"Movement type",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BWART)
Call VerifyGridCellContent("",5,"Article",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MATNR)
Call VerifyGridCellContent("",5,"Document Header Text",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BKTXT)
Call VerifyGridCellContent("",5,"Reason for Movement",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_GRUND)

Call GetGridContent("",0,"Article Document",6,"Document Header Text",DT_MB51_3010_TABLECELL_SINGLE_VALUE_1,"DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_MBLNR_OUTPUT")

Call VerifyGridCellContent("",6,"Posting Date",0,ConvertDate(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BUDAT))
Call VerifyGridCellContent("",6,"Document Header Text",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BKTXT)
Call VerifyGridCellContent("",6,"Reason for Movement",0,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_GRUND)

Call TakeScreenShot() 


Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



