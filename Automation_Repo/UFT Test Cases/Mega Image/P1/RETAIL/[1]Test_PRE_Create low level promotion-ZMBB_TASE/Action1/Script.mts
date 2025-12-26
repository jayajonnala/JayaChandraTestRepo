
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_Create low level promotion-ZMBB
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


gstrTestCaseName = "Test_PRE_promotion-ZMBB"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_PRE_Create low level promotion-ZMBB.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''''''''''''''''''''' TCODE WAK1 '''''''''''''''''''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()   

Call SetTextBox("Promotion type","WAKHD-AKART",0,DT_WAK1_1000_PROMOTION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("Promotion","WAKTD-AKTKT",0,DT_WAK1_0500_PROMOTION,False)
Call ClickButton("Lang\.-dependent description",False)
Call SetTableData("SAPMWAKATEXTE", "Language", 2, "", "", DT_WAK1_5500_TABLECELL_LANGUAGE_1, False)
Call SetTableData("SAPMWAKATEXTE", "Promotion Descr.", 2, "", "", DT_WAK1_5500_TABLECELL_PROMOTION_DESCR_1, False)
Call PressEnter()
Call ClickButton("Back   \(F3\)",False)
Call SetTextBox("On sale from","WAKHD-VKDAB",0,ConvertDate(DT_WAK1_0500_ON_SALE_FROM),False)
Call SetTextBox("to","WAKHD-VKDBI",0,ConvertDate(DT_WAK1_0500_TO),False)
Call SetTextBox("Sales ord\. from","WAKHD-FADAT",0,ConvertDate(DT_WAK1_0500_SALES_ORD_FROM),False)
Call SetTextBox("to","WAKHD-SADAT",0,ConvertDate(DT_WAK1_0500_TO_OCC1),False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Promotion header data   \(Ctrl\+F1\)",False)
Call SetTextBox("Assigned to Promo\.","WAKHD-REF_PROMO",0,DT_WAK1_1000_ASSIGNED_TO_PROMO,False)
Call PressEnter()
Call SetTextBox("Purchasing Org\.","WAKHD-EKORG",0,DT_WAK1_8230_PURCHASING_ORG,False)
Call PressEnter()
Call SelectTab("TAXI_TABSTRIP_KOPFDATEN","Control",False)
Call SetTextBox("Sales Organization","WAKHD-VKORG",0,DT_WAK1_8231_SALES_ORGANIZATION,False)
Call SetTextBox("Distribution Channel","WAKHD-VTWEG",0,DT_WAK1_8231_DISTRIBUTION_CHANNEL,False)
Call SetTextBox("PP determ\. sequence","WAKHD-EKERV",0,"",False)
Call PressEnter()
Call ClickButton("Back   \(F3\)",False)
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Themes",False)
Call SetTableData("SAPMWAKAUEB_ZTHE","Theme",1,"","",DT_WAK1_8214_TABLECELL_THEME_0,False)
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Site groups",False)
Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",1,"","",DT_WAK1_8215_TABLECELL_SITE_GROUP_0,False)
Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",2,"","",DT_WAK1_8215_TABLECELL_SITE_GROUP_1,False)
Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",3,"","",DT_WAK1_8215_TABLECELL_SITE_GROUP_2,False)
Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",4,"","","",False)
Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",5,"","","",False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_WAK1_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Promotion "&DT_WAK1_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
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




