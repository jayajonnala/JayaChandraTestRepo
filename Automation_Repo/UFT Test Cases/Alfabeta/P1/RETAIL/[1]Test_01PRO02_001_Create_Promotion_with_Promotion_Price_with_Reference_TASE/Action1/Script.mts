
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRO02_001_Create_Promotion_with_Promotion_Price_with_Reference
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_01PRO02_001_Create_Promotion_with_Promotion_Price_with_Reference"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRO02_001_Create_Promotion_with_Promotion_Price_with_Reference_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =6
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  WAK1------------------------------------------------


Call SetTextbox("Promotion type","WAKHD-AKART","",DT_WAK1_1000_PROMOTION_TYPE,False)
Call SetTextbox("Ref\. promo\.","RWAKA-KAKTNR","",DT_WAK1_1000_REF_PROMO,False)
Call TakeScreenShot()
Call  ClickButton("Fast entry   \(F5\)",False) 

Call SetTextbox("On sale from","WAKHD-VKDAB","",ConvertDate(DT_WAK1_2600_ON_SALE_FROM),False)
Call SetTextbox("to","WAKHD-VKDBI","",ConvertDate(DT_WAK1_2600_TO),False)
Call SelectCheckbox("RWAKA-KPROMHIERID",0,DT_WAK1_2000_SET_REFERENCE_PROMO_AS_HIGHERLEVEL_PROMO,False)	
Call TakeScreenShot()
Call PressEnter()

Call VerifyStatusBar("Selected data copied from promotion " & DT_WAK1_1000_REF_PROMO)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_PROMO_NUM",Cint(DT_PROMO_NUM)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Promotion","WAKTD-AKTKT","",DT_WAK1_0500_PROMOTION,False)
Call ClickButton("Lang\.-dependent description",False)
Call SetTableData("SAPMWAKATEXTE","Language",2,"","",DT_WAK1_5500_TABLECELL_LANGUAGE_1,False)
Call SetTableData("SAPMWAKATEXTE","Promotion Descr.",2,"","",DT_WAK1_5500_TABLECELL_PROMOTION_DESCR_1,False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False) 
Call  ClickButton("Promotion header data   \(Ctrl\+F1\)",False) 
Call SetTextbox("Purchasing Org\.","WAKHD-EKORG",0,DT_WAK1_8230_PURCHASING_ORG,False)
Call TakeScreenShot()

Call SelectTab("TAXI_TABSTRIP_KOPFDATEN","Control",False)
Call SelectCheckbox("WAKHD-VKKON",0,DT_WAK1_8231_CREATE_SALES_PRICE_COND,False)	
Call SelectCheckbox("WAKHD-NO_KALK",0,DT_WAK1_8231_CREATE_SALES_PRICE_COND,False)
Call SetTextbox("Sales Organization","WAKHD-VKORG","",DT_WAK1_8231_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","WAKHD-VTWEG","",DT_WAK1_8231_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("PP determ\. sequence","WAKHD-EKERV","",DT_WAK1_8231_PP_DETERM_SEQUENCE,False)
Call SetTextbox("Listing Check","WAKHD-VTART",0,DT_LISTING_CHECK,False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False) 
Call PressEnter()
Call SetTableData("SAPMWAKAUEB_SERF","Article","1","","",DT_WAK1_8210_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPMWAKAUEB_SERF","Article","2","","",DT_WAK1_8210_TABLECELL_ARTICLE_1_OCC2,False)
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call TakeScreenShot()
Call GetTableCellData("SAPMWAKAUEB_SERF","Sales Price",1,"","","DT_ORIGINAL_SALES_PRICE_ARTICLE1_OUTPUT",False)
Call GetTableCellData("SAPMWAKAUEB_SERF","Sales Price",2,"","","DT_ORIGINAL_SALES_PRICE_ARTICLE2_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTableData("SAPMWAKAUEB_SERF","Sales Price","1","","",DT_WAK1_8210_TABLECELL_SALES_PRICE_0,False) 
Call SetTableData("SAPMWAKAUEB_SERF","Sales Price","2","","",DT_WAK1_8210_TABLECELL_SALES_PRICE_1,False)
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call TakeScreenShot()


Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Site groups",False)
If DT_SITE_GROUP1<>"" Then
	Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",1,"","",DT_SITE_GROUP1,False)
	Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",2,"","",DT_SITE_GROUP2,False)
	Call SetTableData("SAPMWAKAUEB_ZWRK","Site group",3,"","",DT_SITE_GROUP3,False)
End If
Call TakeScreenShot()
Call PressEnter()

Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call GetStatusBar("item1","DT_PROMOTION_NUMBER_OUTPUT")
Call VerifyStatusBar("Promotion " & DT_PROMOTION_NUMBER_OUTPUT &" created")



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




