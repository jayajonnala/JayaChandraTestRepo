
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_042-Returns to EU Vendors_P1    
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


gstrTestCaseName = "Test_P2P_01_01_042-Returns to EU Vendors_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0156-MD - new vendor with freee goods_P1_Create_PO.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''--------TransactionCode-MB52----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("%_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_MB52_3010_TABLECELL_SINGLE_VALUE_0,True) 
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_MB52_3010_TABLECELL_SINGLE_VALUE_1,True) 
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call SetTextbox("Storage Location","LGORT-LOW","",DT_MB52_1000_STORAGE_LOCATION,False)
Call SetTextbox("Site","WERKS-LOW","",DT_MB52_1000_SITE,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyifGuiLabelExists(trim(DT_MB52_0120_CHECK_TEXT_OF_NO_NAME))
Call VerifyifGuiLabelExists(trim(DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC1))


''''''--------TransactionCode-ME21N----------''''

Call SetTcode(DT_MB52_0120_OKCD)     
Call PressEnter() 

Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call SetComboByKey("MEPO_TOPLINE-BSART", DT_MB52_1105_MEPO_TOPLINEBSART)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_MB52_1105_VENDOR,False)         
Call PressEnter()     
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_MB52_1221_PURCH_ORG,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_MB52_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_MB52_1221_COMPANY_CODE,False)        
Call PressEnter() 
Call TakeScreenShot
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_MB52_1211_TABLECELL_ARTICLE_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_MB52_1211_TABLECELL_ARTICLE_1,False)   
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_MB52_1211_TABLECELL_PO_QUANTITY_0,False)  
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_MB52_1211_TABLECELL_PO_QUANTITY_1,False)  
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_MB52_1211_TABLECELL_SITE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_MB52_1211_TABLECELL_SITE_1,False)
Call PressEnter()
Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call SetTableData("SAPLMEGUITC_1211","Returns Item",1,"","",DT_MB52_1211_TABLECELL_RETURNS_ITEM_0,False)
Call SetTableData("SAPLMEGUITC_1211","Returns Item",2,"","",DT_MB52_1211_TABLECELL_RETURNS_ITEM_1,False)
Call SetTableData("SAPLMEGUITC_1211","Stor. Location","1","","",DT_MB52_1211_TABLECELL_STOR_LOCATION_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Stor. Location","2","","",DT_MB52_1211_TABLECELL_STOR_LOCATION_1,False)
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Purch. Group", "MEPO1222-EKGRP", 0, DT_MB52_1221_CHECK_TEXT_OF_PURCH_GROUP, False)
Call ClickButton("Check   \(Ctrl\+Shift\+F3\)",False)
Call TakeScreenShot
Call VerifyStatusBar(DT_MB52_0014_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot
Call GetStatusBar("item2","DT_MB52_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number "&DT_MB52_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)

''''''--------TransactionCode-ME23N----------''''

Call SetTcode(DT_MB52_0014_OKCD)     
Call PressEnter() 

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_MB52_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot

Call VerifyTableCellContent(1, "Status", "SAPDV70ATC_NAST3", DT_MB52_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_MB52_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1, "Medium", "SAPDV70ATC_NAST3", lcase(DT_MB52_0100_CHECK_VALUE_OF_TABLECELL_MEDIUM_0))

Call LogOff()

Call FinalStatus ()






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


