
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_014-Create low level promotion Special Update prom_P1 
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


gstrTestCaseName = "Test_S2A_PRO_01_014-Create low level promotion Special Update prom_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRO_01_014-Create low level promotion Special Update prom_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario



SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()   

Call SetTextBox("Promotion type","WAKHD-AKART",0,DT_WAK1_1000_PROMOTION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Lang\.-dependent description",False)
'Added below line to update  Language row 1 as EN as it is not coming by default
Call SetTableData("SAPMWAKATEXTE", "Language", 1, "", "", DT_WAK1_5500_TABLECELL_LANGUAGE_0, False)
Call SetTableData("SAPMWAKATEXTE", "Promotion Descr.", 1, "", "", DT_WAK1_5500_TABLECELL_PROMOTION_DESCR_0, False)
Call SetTableData("SAPMWAKATEXTE", "Language", 2, "", "", DT_WAK1_5500_TABLECELL_LANGUAGE_1, False)
Call SetTableData("SAPMWAKATEXTE", "Promotion Descr.", 2, "", "", DT_WAK1_5500_TABLECELL_PROMOTION_DESCR_1, False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Back   \(F3\)",False)
Call SetTextBox("Plnd crrncy","WAKHD-WAELA",0,DT_WAK1_0500_CURRENCY,False)
Call SetTextBox("On sale from","WAKHD-VKDAB",0,ConvertDate(DT_WAK1_0500_ON_SALE_FROM),False)
Call SetTextBox("to","WAKHD-VKDBI",0,ConvertDate(DT_WAK1_0500_TO),False)
Call TakeScreenShot()
Call PressEnter()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Themes",False)
Call TakeScreenShot()
Call SetTableData("SAPMWAKAUEB_ZTHE","Theme",1,"","",DT_WAK1_8214_TABLECELL_THEME_0,False)
Call PressEnter()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Site groups",False)
Call TakeScreenShot()
Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",1,"","",DT_WAK1_8215_TABLECELL_SITE_GROUP_0,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Promotion header data   \(Ctrl\+F1\)",False)
Call TakeScreenShot()
Call SetTextBox("Assigned to Promo.","WAKHD-REF_PROMO",0,DT_WAK1_1000_ASSIGNED_TO_PROMO,False)
Call TakeScreenShot()
Call SetTextBox("Purchasing org.","WAKHD-EKORG",0,DT_WAK1_8230_PURCHASING_ORG,False)
Call SetTextBox("Sales ord. from","WAKHD-FADAT",0,ConvertDate(DT_WAK1_8230_WAKHDZZ_CP1_FROM),False)
Call SetTextBox("to","WAKHD-SADAT",0,ConvertDate(DT_WAK1_8230_WAKHDZZ_CP1_TO),False)
Call PressEnter()
Call TakeScreenShot()
Call GetStatusBar("text","DT_WAK1_8003_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar(DT_WAK1_8003_GET_TEXT_OF_STATUSBAR_OUTPUT)
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_KOPFDATEN","Control",False)
Call TakeScreenShot()
Call SetTextBox("Sales Organization","WAKHD-VKORG",0,DT_WAK1_8231_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","WAKHD-VTWEG",0,DT_WAK1_8231_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("Alloc. Table Type","WAKHD-AUFAR",0,DT_WAK1_8231_ALLOC_TABLE_TYPE,False)
Call SetTextBox("Announcem. category","WAKHD-BEKTY",0,DT_WAK1_8231_ANNOUNCEM_CATEGORY,False)
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_WAK1_1000_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Promotion "&DT_WAK1_1000_GET_TEXT_OF_STATUSBAR_OUTPUT&" created")
Call TakeScreenShot()

Call LogOff()

Call FinalStatus()






'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




