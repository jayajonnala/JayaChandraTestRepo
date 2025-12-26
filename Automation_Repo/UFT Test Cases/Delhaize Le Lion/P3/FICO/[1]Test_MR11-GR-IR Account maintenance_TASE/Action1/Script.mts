
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MR11-GR-IR Account maintenance
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_MR11-GR-IR Account maintenance"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DLL\FICO\TASE_DT_MR11-GR-IR Account maintenance.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call TakeScreenShot
'''SAP Login'''
'
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)
'
Call TakeScreenShot
Call SetTextBox("Company Code","PA_BUKRS",0,DT_MR11_1000_COMPANY_CODE,False)
Call SetTextBox("Reference","PA_XBLNR","",DT_MR11_1000_REFERENCE,False)
Call SetTextBox("Purchasing Document","RA_EBELN-LOW",0,DT_DOCUMENT,False)

Call SetTextBox("Purchase Order Date","RA_BEDAT-LOW",0,DT_MR11_1000_PURCHASE_ORDER_DATE,False)

'Call SetTextBox("Purchase Order Date","RA_BEDAT-LOW",0,ConvertDate(DT_MR11_1000_TO),False)
Call TakeScreenShot
Call SetTextBox("to","RA_BEDAT-HIGH",0,ConvertDate(DT_MR11_1000_TO),False)

Call SetTextBox("Purchasing Document","RA_EBELN-LOW",0,DT_MR11_1000_PURCHASING_DOCUMENT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, "ON", False)
Call TakeScreenShot
Call ClickButton("btn\[36\]",False)

Call GetGridContent("", 0, "Message Text", 1, "Application Area", "M8", "DT_MR11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_CMF_TEXT_OUTPUT")
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call ClickButton("Back   \(F3\)",False)

Call SetTcode(DT_MR11_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)

Call SetTextBox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_MR11_0003_PUR_ORDER,True)

Call PressEnter()
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
' GetGridContent(gridTitle, gridIndex, columnName, rowNumber, refColumn, refFieldVal, dataTableColumnName)
Call GetGridContent("","","Amount",0,"Crcy","EUR",DT_AMOUNT_CHECK)
Call TakeScreenShot
Call SelectTab("ITEM_DETAIL","Account Assignment",False)
Call VerifyTextBoxContent("G/L Account","MEACCT1100-SAKTO","",DT_MR11_1100_CHECK_TEXT_OF_GL_ACCOUNT,False)
Call VerifyTextBoxContent("Cost Center","COBL-KOSTL","",DT_MR11_1101_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)

Call SetTcode(DT_MR11_0100_OKCD_OCC1)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Account Number","RACCT-LOW",0,DT_MR11_1000_ACCOUNT_NUMBER,False)

Call SetTextBox("Company Code","RBUKRS-LOW",0,DT_MR11_1000_COMPANY_CODE_OCC1,False)

Call SetTextBox("Fiscal Year","RYEAR",0,DT_MR11_1000_FISCAL_YEAR,False)

Call ClickButton("Free Selections   \(Ctrl\+F1\)",False)
Call TakeScreenShot
Call SetTextBox("Cost Center","%%DYN002-LOW",0,DT_MR11_0100_COST_CENTER,False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()
'Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButton("Find",True)
'Call TakeScreenShot
'Call SetTextBox("Find","GD_SEARCHSTR","",DT_MR11_0841_SEARCH_TERM_OCC1,True)
'Call ClickButton("Continue   \(Enter\)",True)
'Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
'Call SetTextBox("Purchasing Document","%%DYN001-LOW","",DT_MR11_3010_TABLECELL_SINGLE_VALUE_0,True)
'Call ClickButton("Execute   \(Enter\)",True)

Call LogOff()
Call FinalStatus ()
'''''''
'''''''
'''''''''INPUT''
'''''''' SetTextbox(textboxAttachedText, textboxName, textboxIndex, textboxValue, blnIsItPopup)
'''''''call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,false)
'''''''call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,false)
'''''''call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,false)
'''''''call PressEnter
'''''''
''''''''''create vendor : Address
''''''''SetCombo(attachedTextOrComboName,comboValue)
'''''''call SetCombo("Title",DT_XK01_0301_TITLE)
'''''''call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,false)
'''''''call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,false)
'''''''call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,false)
'''''''call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,false)
'''''''call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,false)
'''''''call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,false)
'''''''call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,false)
'''''''call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XK01_0301_TELEPHONE,false)
''''''''call SetCombo("ADDR1_DATA-LANGU",DT_XK01_0301_LANGUAGE)
'''''''call ClickButton("G_ICON_SMTP",false)
'''''''
''''''''''Maintain internet mail addresses'''
'''''''' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
'''''''call SetTableData("SAPLSZA6T_CONTROL6","E-Mail Address",1,"","",DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_0,false)
'''''''call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_0,false)
'''''''call ClickButton("btn\[13\]",false)
''''''''''2nd
'''''''call SetTableData("SAPLSZA6T_CONTROL6","E-Mail Address",1,"","",DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_1,false)
'''''''call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_1,false)
'''''''call ClickButton("btn\[13\]",false)
''''''''''3rd
'''''''call SetTableData("SAPLSZA6T_CONTROL6","E-Mail Address",1,"","",DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_2,false)
'''''''call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_2,false)
'''''''
'''''''call ClickButton("btn\[0\]",false)
'''''''
''''''''''VAT Reg No
''''''''call SetTextbox("VAT Reg\. No\.","LFA1-STCEG","",DT_XK01_0120_VAT_REG_NO,false)  '''No need to give VAT No
'''''''call PressEnter
'''''''call PressEnter
'''''''call PressEnter
'''''''' ClickCellTable(tableName, columnName, rowNumber, columnNameRef, tableValRef, blnIsItPopup)
'''''''call ClickCellTable("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","IBAN",1,"IBAN","123",False)
'''''''call ClickButton("SWITCH",false)
'''''''call SetTextbox("IBAN","IBAN00","",DT_XK01_0200_IBAN,false)
'''''''call SetTextbox("SWIFT/BIC","BNKA-SWIFT","",DT_XK01_0200_SWIFT_CODE,false)
'''''''call ClickButton("btn\[7\]",false)
'''''''call ClickButton("btn\[0\]",false)
'''''''call ClickButton("btn\[0\]",false)  ''additional steps
'''''''call SetTableData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Acct Holder",1,"","",DT_XK01_0130_TABLECELL_ACCT_HOLDER_0,false)
'''''''call SetTableData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","BnkT",1,"","",DT_XK01_0130_TABLECELL_BNKT_0,false)
'''''''
'''''''call PressEnter
'''''''call PressEnter
'''''''call PressEnter
'''''''''''CREATE VENDOR ACCOUNTING INFORMATION'''
'''''''call SetTextbox("Recon\. account","LFB1-AKONT","",DT_XK01_0210_RECON_ACCOUNT,false)
'''''''call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,false)
'''''''call PressEnter
'''''''''''CREATE VENDOR : PAYMENT TRANSACTION ACCOUNITING''''
'''''''
'''''''
'''''''
'''''''call SetTextbox("Payt Terms","LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,false)
'''''''' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
'''''''
'''''''call SelectCheckbox("LFB1-REPRF","1","ON",false)
'''''''
'''''''call SetTextbox("Cr memo terms","LFB1-GUZTE","",DT_XK01_0215_CR_MEMO_TERMS,false)
'''''''call SetTextbox("Tolerance group","LFB1-TOGRR","",DT_XK01_0215_TOLERANCE_GROUP,false)
'''''''call SetTextbox("Assign\.Grp","LFB1-ASSIGN_TEST","",DT_XK01_0215_ASSIGNGRP,false)
'''''''call SetTextbox("Payment methods","LFB1-ZWELS","","S",false)
'''''''call PressEnter
'''''''wait(3)
'''''''call PressEnter
'''''''wait(3)
'''''''call PressEnter
'''''''wait(3)
''''''''''Create vendor : purchasing data'''
'''''''
'''''''call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,false)
'''''''call SetTextbox("Terms of paymnt","LFM1-ZTERM","",DT_XK01_0310_TERMS_OF_PAYMNT,false)
'''''''call SelectCheckbox("LFM1-BLIND","1","ON",false)
'''''''call SelectCheckbox("LFM1-AGREL","1","ON",false)
'''''''''''environment/classification'''
'''''''
'''''''call SelectMenuBar("Environment;Classification")
'''''''call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class","1","","",DT_XK01_1600_TABLECELL_CLASS_0,false)
'''''''call PressEnter
'''''''
'''''''call SetTableData("SAPLCTMSCHARS_S","Value","4","","",DT_XK01_5100_TABLECELL_VALUE_4,false)
'''''''call PressEnter
'''''''' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
'''''''call ClickButtonIfExist("btn\[3\]",false)
'''''''call ClickButtonIfExist("btn\[11\]",false)
'''''''' GetStatusBar(itemNo, dataTableColumnName)
'''''''call GetStatusBar("item1","DT_Output_Vendor")
'''''''Call VerifyStatusBar("created ")

