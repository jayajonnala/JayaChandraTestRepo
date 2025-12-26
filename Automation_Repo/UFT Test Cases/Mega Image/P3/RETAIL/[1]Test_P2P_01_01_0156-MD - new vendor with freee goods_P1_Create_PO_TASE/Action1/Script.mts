
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0156-MD - new vendor with freee goods_P1_Create_PO     
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

gstrTestCaseName = "Test_P2P_01_01_0156-MD_P1_Create_PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0156-MD - new vendor with freee goods_P1_Create_PO.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''''''--------TransactionCode-MBN1----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Discount type","N000-KSCHL","",DT_MBN3_0100_DISCOUNT_TYPE,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectRadioButton("RV130-SELKZ","Info Record per Order Unit", True)
Call ClickButtonIfExist("Choose   \(Enter\)",True)
'''Call SetTextbox("Vendor","F001","",DT_MBN3_1000_VENDOR,False)
Call SetTextboxNoLabel("KOMG-LIFNR","",DT_MBN3_1000_VENDOR,False)
Call SetTextbox("Article","KOMG-MATNR","",DT_MBN3_1000_ARTICLE,False)
Call SetTextbox("Purch. Organization","KOMG-EKORG","",DT_MBN3_1000_PURCH_ORGANIZATION,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call SetTableData("SAPMV13NTCTRL_FAST_ENTRY", "Infotype", 1, "<NA>", "<NA>", DT_INFOTYPE, False)
Call SetTableData("SAPMV13NTCTRL_FAST_ENTRY", "Order Unit", 1, "<NA>", "<NA>", DT_UNIT, False)
Call SetTableData("SAPMV13NTCTRL_FAST_ENTRY", "Min. qty", 1, "<NA>", "<NA>", DT_MBN1_1002_CHECK_TEXT_OF_TABLECELL_MIN_QTY_0, False)
Call SetTableData("SAPMV13NTCTRL_FAST_ENTRY", "From", 1, "<NA>", "<NA>", DT_MBN1_1002_CHECK_TEXT_OF_TABLECELL_FROM_0, False)
Call SetTableData("SAPMV13NTCTRL_FAST_ENTRY", "UnitFG", 1, "<NA>", "<NA>", DT_UNIT, False)
Call SetTableData("SAPMV13NTCTRL_FAST_ENTRY", "are free goo", 1, "<NA>", "<NA>", DT_FREE_GOODS, False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Confirm   \(Enter\)", False)
Call TakeScreenShot
Call VerifyStatusBar("Condition records for Free goods saved")



''''''--------TransactionCode-MBN3----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Discount type","N000-KSCHL","",DT_MBN3_0100_DISCOUNT_TYPE,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectRadioButton("RV130-SELKZ","Info Record per Order Unit", True)
Call ClickButtonIfExist("Choose   \(Enter\)",True)
Call SetTextbox("Purch. Organization","F003","",DT_MBN3_1000_PURCH_ORGANIZATION,False)
'''Call SetTextbox("Vendor","F001","",DT_MBN3_1000_VENDOR,False)
Call SetTextboxNoLabel("F001","",DT_MBN3_1000_VENDOR,False)
Call SetTextbox("Article","F002","",DT_MBN3_1000_ARTICLE,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call GetTableCellData("SAPMV13NTCTRL_FAST_ENTRY", "Min. qty", 1, "<NA>", "<NA>", "DT_MBN3_1002_CHECK_TEXT_OF_TABLECELL_MIN_QTY_0_OUTPUT", False)
Call GetTableCellData("SAPMV13NTCTRL_FAST_ENTRY", "From", 1, "<NA>", "<NA>", "DT_MBN3_1002_CHECK_TEXT_OF_TABLECELL_FROM_0_OUTPUT", False)
'
''''''--------TransactionCode-ME21N----------''''

Call SetTcode(DT_MBN3_1002_OKCD)     
Call PressEnter() 

Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
''Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_MBN3_1105_VENDOR,False)         
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_MBN3_1105_VENDOR,False)    
Call PressEnter()     
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_MBN3_1221_PURCH_ORG,False)     
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_MBN3_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_MBN3_1221_COMPANY_CODE,False)        
Call PressEnter() 
Call TakeScreenShot
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_MBN3_1211_TABLECELL_ARTICLE_0,False)     
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_MBN3_1211_TABLECELL_PO_QUANTITY_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_MBN3_1211_TABLECELL_DELIV_DATE_0),False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_MBN3_1211_TABLECELL_SITE_0,False) 
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call VerifyTableCellContent(2, "Net Price", "SAPLMEGUITC_1211", DT_MBN3_1211_CHECK_TEXT_OF_TABLECELL_NET_PRICE_1)
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot
Call GetStatusBar("item2","DT_MBN3_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Standard PO created under the number "&DT_MBN3_0014_CHECK_TEXT_OF_STATUSBAR_OUTPUT)


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


