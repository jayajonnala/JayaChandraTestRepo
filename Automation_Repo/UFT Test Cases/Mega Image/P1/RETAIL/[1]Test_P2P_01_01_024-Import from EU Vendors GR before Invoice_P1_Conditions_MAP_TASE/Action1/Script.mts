

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_024-Import from EU Vendors GR before Invoice_P1_Conditions_MAP    
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


gstrTestCaseName = "Test_P2P_01_01_024- Invoice_P1_Conditions_MAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_024-Import from EU Vendors GR before Invoice_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",false)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME22N_0003_PUR_ORDER_OCC1,True)
Call ClickButton("Other Document   \(Enter\)",True)
'Call ClickButtonIfExist("Display/Change   \(F7\)",False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call SelectTab("HEADER_DETAIL","Conditions",False)

Call ClickButton("Insert Row",False)
wait 5
Call  SetTableData("SAPLV69ATCTRL_KONDITIONEN","CnTy",23,"","",DT_ME22N_6201_TABLECELL_CNTY_1,False)
Call  SetTableData("SAPLV69ATCTRL_KONDITIONEN","Amount",23,"","",DT_ME22N_6201_TABLECELL_AMOUNT_1,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Collapse Header Ctrl\+F5",False)
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Conditions",False)
Call VerifyTableCellContent(10,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_ME22N_6201_CHECK_TEXT_OF_TABLECELL_CNTY_4)
Call VerifyTableCellContent(10,"Non-Active","SAPLV69ATCTRL_KONDITIONEN","S_LEDG")
Call TakeScreenShot
Call ClickButton("Next item",False)
Call VerifyTableCellContent(10,"CnTy","SAPLV69ATCTRL_KONDITIONEN",DT_ME22N_6201_CHECK_TEXT_OF_TABLECELL_NAME_4)
Call VerifyTableCellContent(10,"Non-Active","SAPLV69ATCTRL_KONDITIONEN",DT_ME22N_6201_CHECK_TOOLTIP_OF_TABLECELL_NONACTIVE_4_OCC1)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call TakeScreenShot
Call VerifyStatusBar("Standard PO Retail " &DT_ME22N_0003_PUR_ORDER_OCC1& " changed")
Call SetTcode(DT_ME22N_0014_OKCD)     
Call PressEnter() 
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",false)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME22N_0003_PUR_ORDER_OCC1,True)
Call ClickButton("Other Document   \(Enter\)",True)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call SelectTab("HEADER_DETAIL","Release strategy",False)
Call ClickCellGuiGrid("",0,"Release options",1,"","",False)
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_ME22N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_ME22N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(5,"Output Type","SAPDV70ATC_NAST3",DT_ME22N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)
Call VerifyTableCellContent(5,"Status","SAPDV70ATC_NAST3",DT_ME22N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_3)
Call TakeScreenShot
Call SetTcode(DT_ME22N_0201_OKCD)     
Call PressEnter()  


Call SetTextbox("Article","RMMW1-MATNR","",DT_ME22N_0100_ARTICLE,False)
Call SetTextbox("Distr\. center","RMMW1-VZWRK","",DT_ME22N_0100_DISTR_CENTER,False)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_ME22N_0100_PURCHASING_ORG,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",5,False)
Call PressEnter()  
Call TakeScreenShot
Call SelectTab("TABSPR1"," Logistics: DC",False)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call GetTextboxValue("MBEW-VERPR",9,"DT_ME22N_2802_CHECK_TEXT_OF_MOVING_PRICE_OUTPUT",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call SetTextbox("Article","RMMW1-MATNR","",DT_ME22N_0100_ARTICLE_OCC1,False)
Call SetTextbox("Distr\. center","RMMW1-VZWRK","",DT_ME22N_0100_DISTR_CENTER_OCC1,False)
Call SetTextbox("Purchasing Org\.","RMMW1-EKORG","",DT_ME22N_0100_PURCHASING_ORG_OCC1,False)
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",5,False)
Call PressEnter()  
Call TakeScreenShot
Call SelectTab("TABSPR1"," Logistics: DC",False)
Call TakeScreenShot
Call ClickButton("Accounting",False)
Call GetTextboxValue("MBEW-VERPR",9,"DT_ME22N_2802_CHECK_TEXT_OF_MOVING_PRICE_OCC1_OUTPUT",False)

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


