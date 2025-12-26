
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03-02-01-01-02-Create Cust Retail_v2
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

gstrTestCaseName = "Test_03-02-01-01-02-Create Cust Retail_v2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

''INPUT''
call ClickButton("KONTENGRUPPE_INFO",false)  ''''uncomment
call SetTextbox("Account group","\*T077D-KTOKD","",DT_XD01_1005_ACCOUNT_GROUP,true)  ''its not working
'call ClickButton("\*T077D-KTOKD",false)  '''' NOT req


call ClickButton("btn\[0\]",false)

''check'''
'call SelectRowGuiTable("SAPMF02DTCTRL_KONTENGRUPPEN","Group","ZCXT",true)
call SelectCellGuiTable("SAPMF02DTCTRL_KONTENGRUPPEN","Group","Group",DT_XD01_1005_ACCOUNT_GROUP,False)
'call SetTableData("SAPMF02DTCTRL_KONTENGRUPPEN","Group",1,"","","",false)
''''click on first table'''

call ClickButton("btn\[0\]",true)
''''click on first table'''
'call ClickButton("btn\[0\]",false)


call SetTextbox("Company code","RF02D-BUKRS","",DT_XD01_7100_COMPANY_CODE,false)
call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD01_7100_SALES_ORGANIZATION,false)
call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD01_7100_DISTRIBUTION_CHANNEL,false)
call SetTextbox("Division","RF02D-SPART","",DT_XD01_7100_DIVISION,false)
call PressEnter
'''
''execution 
call SetCombo("Title",DT_XD01_0301_TITLE)
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_03-02-01-01-02-Create Customer in Retail.xls"
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",2)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XD01_0301_NAME,false)
call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XD01_0301_SEARCH_TERM_12,false)
call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XD01_0301_STREETHOUSE_NUMBER,false)
call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XD01_0301_STREETHOUSE_NUMBER_OCC1,false)
call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XD01_0301_POSTAL_CODECITY,false)
call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XD01_0301_POSTAL_CODECITY_OCC1,false)
call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XD01_0301_COUNTRY,false)
call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XD01_0301_TELEPHONE,false)
'call SetCombo("ADDR1_DATA-LANGU",DT_XD01_0301_LANGUAGE)
call SetTextbox("E-Mail","SZA1_D0100-SMTP_ADDR","",DT_XD01_0301_EMAIL,false)
Call SelectTab("TABSTRIP100","Control Data",False)

call SetTextboxNoLabel("KNA1-LIFNR","",DT_XD01_7121_VENDOR,false)
call SetTextbox("VAT Reg\. No\.","KNA1-STCEG","",DT_XD01_7122_VAT_REG_NO,false)
call PressEnter
call PressEnter
Call SelectTab("TABSTRIP100","Payment Transactions",False)

call ClickButton("btn\[26\]",false)
call SetTextbox("Recon\. account","KNB1-AKONT","",DT_XD01_7211_RECON_ACCOUNT,false)
call SetTextbox("Sort key","KNB1-ZUAWA","",DT_XD01_7211_SORT_KEY,false)
call SetTextbox("Cash mgmt group","KNB1-FDGRV","",DT_XD01_7211_CASH_MGMT_GROUP,false)
Call SelectTab("TABSTRIP100","Payment Transactions",False)

call SetTextbox("Terms of payment","KNB1-ZTERM","",DT_XD01_7215_TERMS_OF_PAYMENT,false)
' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
'call SelectCheckbox("KNB1-XVERR",3,DT_XD01_7216_CLEARING_WITH_VENDOR,false)
call SelectCheckbox("KNB1-XZVER",1,DT_XD01_7215_PAYMENT_HISTORY_RECORD,false)

Call SelectTab("TABSTRIP100","Correspondence",False)
call SetTextbox("Dunn\.Procedure","KNB5-MAHNA","",DT_XD01_7220_DUNNPROCEDURE,false)
call SetTextbox("Dunning clerk","KNB5-BUSAB","",DT_XD01_7220_DUNNING_CLERK,false)
call ClickButton("btn\[27\]",false)
call SetTextbox("Cust\.pric\.proc\.","KNVV-KALKS","",DT_XD01_7311_CUSTPRICPROC,false)
call SelectCheckbox("KNVV-AGREL",1,DT_XD01_7312_RELEVANT_FOR_AGENCY_BUSINESS,false)
Call SelectTab("TABSTRIP100","Shipping",False)
call SetTextbox("Shipping Conditions","KNVV-VSBED","",DT_XD01_7315_SHIPPING_CONDITIONS,false)
Call SelectTab("TABSTRIP100","Billing Documents",False)
call SetTextbox("Terms of payment","KNVV-ZTERM","",DT_XD01_7321_TERMS_OF_PAYMENT,false)
call SetTextbox("Acct assgmt group","KNVV-KTGRD","",DT_XD01_7322_ACCT_ASSGMT_GROUP,false)
'SetTableData(tableName,columnName,rowNumber,refColumnName,refCellValue,cellValue,blnIsItPopup)

call SetTableData("SAPMF02DTCTRL_STEUERN","Tax classification",1,"","",DT_XD01_7350_TABLECELL_TAX_CLASSIFICATION_0,false)

call PressEnter
Call SelectMenuBar("Extras;Classification")

call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class",1,"","",DT_XD01_1600_TABLECELL_CLASS_0,false)

call PressEnter
call ClickButton("Back   \(F3\)",false)
call ClickButton("Save   \(Ctrl\+S\)",false)
call GetStatusBar("item1","DT_Statusbar_Output")
Call verifystatusbar("Company Code BE10")
Call LogOff()
Call FinalStatus ()
'''control data

