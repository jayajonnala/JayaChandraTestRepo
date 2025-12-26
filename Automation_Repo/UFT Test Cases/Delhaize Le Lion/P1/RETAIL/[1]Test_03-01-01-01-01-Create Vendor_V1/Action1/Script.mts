
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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


gstrTestCaseName = "Test_03-01-01-01-01-Cre Ven_V1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)
Call TakeScreenShot

''INPUT''
' SetTextbox(textboxAttachedText, textboxName, textboxIndex, textboxValue, blnIsItPopup)
call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,false)
call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,false)
call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,false)
Call TakeScreenShot
call PressEnter

'''create vendor : Address
'SetCombo(attachedTextOrComboName,comboValue)
call SetCombo("Title",DT_XK01_0301_TITLE)
call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,false)
call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,false)
call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,false)
call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,false)
call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,false)
call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,false)
call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,false)
call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XK01_0301_TELEPHONE,false)
Call TakeScreenShot
'call SetCombo("ADDR1_DATA-LANGU",DT_XK01_0301_LANGUAGE)
call ClickButton("G_ICON_SMTP",false)

'''Maintain internet mail addresses'''
' SetTableData(tableName, columnName, rowNumber, refColumnName, refCellValue, cellValue, blnIsItPopup)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
call SetTableData("SAPLSZA6T_CONTROL6","E-Mail Address",1,"","",DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_0,false)
call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_0,false)
call ClickButton("btn\[13\]",false)
Call TakeScreenShot
'''2nd
call SetTableData("SAPLSZA6T_CONTROL6","E-Mail Address",1,"","",DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_1,false)
call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_1,false)
call ClickButton("btn\[13\]",false)
Call TakeScreenShot
'''3rd
call SetTableData("SAPLSZA6T_CONTROL6","E-Mail Address",1,"","",DT_XK01_0600_TABLECELL_EMAIL_ADDRESS_2,false)
call SetTableData("SAPLSZA6T_CONTROL6","Notes",1,"","",DT_XK01_0600_TABLECELL_NOTES_2,false)
Call TakeScreenShot
call ClickButton("btn\[0\]",false)

'''VAT Reg No
'call SetTextbox("VAT Reg\. No\.","LFA1-STCEG","",DT_XK01_0120_VAT_REG_NO,false)  '''No need to give VAT No
call PressEnter
call PressEnter
call PressEnter
' ClickCellTable(tableName, columnName, rowNumber, columnNameRef, tableValRef, blnIsItPopup)
call ClickCellTable("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","IBAN",1,"IBAN","123",False)
Call TakeScreenShot
call ClickButton("SWITCH",false)
Call TakeScreenShot
call SetTextbox("IBAN","IBAN00","",DT_XK01_0200_IBAN,false)
call SetTextbox("SWIFT/BIC","BNKA-SWIFT","",DT_XK01_0200_SWIFT_CODE,false)
call ClickButton("btn\[7\]",false)
Call VerifyStatusBarMessageType("I")
Call TakeScreenShot
call ClickButton("btn\[0\]",false)
call ClickButton("btn\[0\]",false)  ''additional steps
Call TakeScreenShot
call SetTableData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","Acct Holder",1,"","",DT_XK01_0130_TABLECELL_ACCT_HOLDER_0,false)
call SetTableData("SAPMF02KTCTRL_ZAHLUNGSVERKEHR","BnkT",1,"","",DT_XK01_0130_TABLECELL_BNKT_0,false)
Call TakeScreenShot
call PressEnter
call PressEnter
call PressEnter
''''CREATE VENDOR ACCOUNTING INFORMATION'''
call SetTextbox("Recon\. account","LFB1-AKONT","",DT_XK01_0210_RECON_ACCOUNT,false)
call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,false)
Call TakeScreenShot
call PressEnter
''''CREATE VENDOR : PAYMENT TRANSACTION ACCOUNITING''''

call SetTextboxNoLabel("LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,false)
' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
'This function SetTextboxNoLabel is used for vendor to supplier change.
call SelectCheckbox("LFB1-REPRF","1","ON",false)
Call TakeScreenShot
call SetTextbox("Cr memo terms","LFB1-GUZTE","",DT_XK01_0215_CR_MEMO_TERMS,false)
call SetTextbox("Tolerance group","LFB1-TOGRR","",DT_XK01_0215_TOLERANCE_GROUP,false)
call SetTextbox("Assign\.Grp","LFB1-ASSIGN_TEST","",DT_XK01_0215_ASSIGNGRP,false)
call SetTextbox("Payment methods","LFB1-ZWELS","","S",false)
Call TakeScreenShot
call PressEnter
Call TakeScreenShot
wait(3)
call PressEnter
wait(3)
call PressEnter
wait(3)
'''Create vendor : purchasing data'''

call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,false)
call SetTextbox("Terms of paymnt","LFM1-ZTERM","",DT_XK01_0310_TERMS_OF_PAYMNT,false)
call SelectCheckbox("LFM1-BLIND","1","ON",false)
call SelectCheckbox("LFM1-AGREL","1","ON",false)
Call TakeScreenShot
''''environment/classification'''

call SelectMenuBar("Environment;Classification")
call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class","1","","",DT_XK01_1600_TABLECELL_CLASS_0,false)
call PressEnter
Call TakeScreenShot
call SetTableData("SAPLCTMSCHARS_S","Value","5","","",DT_XK01_5100_TABLECELL_VALUE_4,false)
call PressEnter
Call TakeScreenShot
' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
call ClickButtonIfExist("btn\[3\]",false)
Call TakeScreenShot
call ClickButtonIfExist("btn\[11\]",false)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
' GetStatusBar(itemNo, dataTableColumnName)
call GetStatusBar("item1","DT_Output_Vendor")
Call WriteRunTimeDataToExcelGlobalSheet("DT_Output_Vendor",DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_XK01_0100_CHECK_TEXT_OF_STATUSBAR)

'
'strval = len(val)
'msgbox Mid(strval,1,4)
'msgbox Mid(strval,5,4)
'msgbox Mid(strval,9,4)
'msgbox strval
'For i = 1 To len((val)) Step 4
'	strval1 =  val(i)
'	msgbox strval1
'Next
'
Call LogOff()
Call FinalStatus ()
