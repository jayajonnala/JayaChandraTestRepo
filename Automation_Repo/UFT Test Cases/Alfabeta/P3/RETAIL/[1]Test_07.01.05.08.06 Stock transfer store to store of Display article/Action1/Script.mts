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

gstrTestCaseName = "Test_07.01.05.08.06 Display article"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.01.05.08.06 Stock transfer store to store of Display article_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'//-----------------------------------MIGO -----------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenshot()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''Call SetCombo("GODYNPRO-ACTION","Transfer Posting")
'''Call SetCombo("GODYNPRO-REFDOC","Other")
Call SetComboByKey("GODYNPRO-ACTION","A08")
Call SetComboByKey("GODYNPRO-REFDOC","R10")
Call TakeScreenshot()

Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
'Call SetTextbox("TF trfr plnt to plnt","GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False) 
Call TakeScreenshot()

Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_POSTING_DATE),False)

call ClickButtonIfExist("Close Detail Data",false)
Call TakeScreenshot()

call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",1,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Mat. Short Text",2,"","",DT_MIGO_0200_TABLECELL_ART_SHORT_TEXT_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",1,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Qty in UnE",2,"","",DT_MIGO_0200_TABLECELL_QTY_IN_UNE_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",1,"","",DT_MIGO_0200_TABLECELL_SLOC_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc",2,"","",DT_MIGO_0200_TABLECELL_SLOC_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","Plant Trfr Pstg",1,"","",DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_0,False)
call SetTableData("SAPLMIGOTV_GOITEM","Plant Trfr Pstg",2,"","",DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_1,False)
call SetTableData("SAPLMIGOTV_GOITEM","Plnt",1,"","",DT_MIGO_0200_TABLECELL_SITE_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","Plnt",2,"","",DT_MIGO_0200_TABLECELL_SITE_1,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",1,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_0,false)
call SetTableData("SAPLMIGOTV_GOITEM","SLoc Transfer Pstg",2,"","",DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_1,false)
Call TakeScreenshot()
Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextbox("Article Slip","GOHEAD-MTSNR","",DT_MIGO_0112_ARTICLE_SLIP,False) 
Call PressEnter()
Call TakeScreenshot()

call ClickButton("Post Document   \(Shift\+F11\)",fALSE)
Call TakeScreenshot()
call ClickButtonIfExist("Continue   \(Enter\)",fALSE)
Call TakeScreenshot()
call GetStatusBar("item1","DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR")
Call WriteRunTimeDataToExcelGlobalSheet("DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR",DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)


Call LogOff()
Call FinalStatus ()








''DT_MIGO_0200_TABLECELL_SLOC_TRANSFER_PSTG_0
'
''Call PressEnter()
''call SetTableData("SAPLMIGOTV_GOITEM","Plant Trfr Pstg",1,"","","ΨΥΧΙΚΟ",false)
''call SetTableData("SAPLMIGOTV_GOITEM","Plant Trfr Pstg",2,"","","ΨΥΧΙΚΟ",false)
''DT_MIGO_0200_TABLECELL_SITE_TRFR_PSTG_0
'Call PressEnter()
'IF SAPGuiSession("Session_2").SAPGuiWindow("Transfer Posting Other").SAPGuiLabel("G001").EXIST THEN
'    SAPGuiSession("Session_2").SAPGuiWindow("Transfer Posting Other").SAPGuiLabel("G001").SetFocus
'    
'    
'End If
''call SetFocusGuiLabel("G001",11,56,true)
'call ClickButtonIfExist("Copy   \(Enter\)",TRUE)
'SAPGuiSession("Session_2").SAPGuiWindow("Transfer Posting Other").SAPGuiLabel("G001").SetFocus
'' SetFocusGuiLabel(labelContent, xCord, yCord, blnIsItPopup)
'call ClickButtonIfExist("Copy   \(Enter\)",TRUE)
'





'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1        This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2        Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3        User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4        Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5        If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

