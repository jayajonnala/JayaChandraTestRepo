	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_07.01.05.08.010  same delivery note"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.01.05.08.010 Negative_ Store to Store with same delivery note_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)

Call SetCombo("GODYNPRO-ACTION","Transfer Posting")
Call SetCombo("GODYNPRO-REFDOC","Other")
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_PO_NUMBER_OUTPUT,False)

Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)

Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False) 
call ClickButtonIfExist("Close Detail Data",false)

' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
'call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,false) 'Changed in RGB
call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",2,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",1,"","",DT_MIGO_0200_TABLECELL_SLOC_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","SLoc",2,"","",DT_MIGO_0200_TABLECELL_SLOC_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",1,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",2,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",2,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_1,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Art. Short Text",2,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_1,false) 'Changed in RGB
call SetTableData("SAPLMIGOTV_GOITEM","Plnt",1,"","",DT_MIGO_0200_TABLECELL_SITE_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Plnt",2,"","",DT_MIGO_0200_TABLECELL_SITE_1,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Site",1,"","",DT_MIGO_0200_TABLECELL_SITE_0,false) 'Changed in RGB
'call SetTableData("SAPLMIGOTV_GOITEM","Site",2,"","",DT_MIGO_0200_TABLECELL_SITE_1,false) 'Changed in RGB
call SetTableData("SAPLMIGOTV_GOITEM","Plant Trfr Pstg",1,"","",DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_0,false)
'call SetTableData("SAPLMIGOTV_GOITEM","Site Trfr Pstg",2,"","",DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_1,false)
'Call PressEnter()
Call PressEnter()
'Call ClickLabel("Werk 0001",1,true)



' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0112_DOCUMENT_DATE),False)
Call VerifyTextBoxContent("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0112_POSTING_DATE),False)
Call VerifyTextBoxContent("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False)
call ClickButtonIfExist("Check Entries   \(F7\)",False)

Call CheckifGuiLabelExists(DT_ERROR_MESSAGE)
call ClickButtonIfExist("Continue   \(Enter\)",fALSE)


'call ClickButtonIfExist("Copy   \(Enter\)",fALSE)
'call ClickButton("Post Document   \(Shift\+F11\)",fALSE)
'call ClickButtonIfExist("Continue   \(Enter\)",fALSE)


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


