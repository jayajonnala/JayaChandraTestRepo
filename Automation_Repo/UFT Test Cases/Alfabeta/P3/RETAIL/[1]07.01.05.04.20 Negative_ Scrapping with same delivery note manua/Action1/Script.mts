	

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

gstrTestCaseName = "07.01.05.04.20 Negative same delivery note "
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.01.05.04.20 Negative_ Scrapping with same delivery note manua_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
Call SetComboByKey("GODYNPRO-ACTION",DT_ACTION)
Call SetComboByKey("GODYNPRO-REFDOC",DT_DOCREF)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call TakeScreenShot
' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MIGO_0112_ARTICLE_SLIP",Cint(DT_MIGO_0112_ARTICLE_SLIP)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False) 
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_POSTING_DATE),False) 
call ClickButtonIfExist("Close Detail Data",false)
Call TakeScreenShot
' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",2,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",3,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_2,false)

call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",2,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",3,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_2,false)
Call TakeScreenShot
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",1,"","",DT_MIGO_0200_TABLECELL_SLOC_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",2,"","",DT_MIGO_0200_TABLECELL_SLOC_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",3,"","",DT_MIGO_0200_TABLECELL_SLOC_2,false)

call SetTableData("SAPLMIGOTV_GOITEM","Plnt",1,"","",DT_MIGO_0200_TABLECELL_SITE_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Plnt",2,"","",DT_MIGO_0200_TABLECELL_SITE_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Plnt",3,"","",DT_MIGO_0200_TABLECELL_SITE_2,false)
Call TakeScreenShot
wait 2
call SetTableData("SAPLMIGOTV_GOITEM","Reason for Mvmt",1,"","",DT_MIGO_0200_TABLECELL_REASON_FOR_MVMT_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Reason for Mvmt",2,"","",DT_MIGO_0200_TABLECELL_REASON_FOR_MVMT_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Reason for Mvmt",3,"","",DT_MIGO_0200_TABLECELL_REASON_FOR_MVMT_2,false)


''call SetTableData("SAPLMIGOTV_GOITEM","Site Trfr Pstg",1,"","",DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_0,false)
''call SetTableData("SAPLMIGOTV_GOITEM","Site Trfr Pstg",2,"","",DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_1,false)
''call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",1,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_0,false)
''call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",2,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_1,false)
''call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",3,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_2,false)

Call PressEnter()
call ClickButton("Check Entries   \(F7\)",False)
Call TakeScreenShot
call ClickButton("Continue   \(Enter\)",False)
call ClickButton("Post Document   \(Shift\+F11\)",False)
' GetStatusBar(itemNo, dataTableColumnName)
Call TakeScreenShot
Call GetStatusBar("item1","DT_ARTICLE_DOCNUM_OUTPUT")
Call VerifyStatusBar("Article document "&DT_ARTICLE_DOCNUM_OUTPUT&" posted")
Call TakeScreenShot
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_ARTICLE_DOCNUM_OUTPUT",DT_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0_OCC1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_0_OCC1,false)
Call TakeScreenShot
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",1,"","",DT_MIGO_0200_TABLECELL_SLOC_0_OCC1,false)
Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP_OCC1,False) 
call SetTableData("SAPLMIGOTV_GOITEM","Plnt",1,"","",DT_MIGO_0200_TABLECELL_SITE_0_OCC1,false)
Call TakeScreenShot
Call PressEnter()
call SetTableData("SAPLMIGOTV_GOITEM","Reason for Mvmt",1,"","",DT_MIGO_0200_TABLECELL_REASON_FOR_MVMT_0_OCC1,false)
''call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",1,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_0_OCC1,false)
Call PressEnter()
Call ClickButtonIfExist("Check Entries   \(F7\)", False)
Call TakeScreenShot
wait 2
Call VerifyifGuiLabelExists (DT_MIGO_0120_CHECK_TEXT_OF_DELIVERY_NOTE)
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


