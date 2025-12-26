

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_AB_004 Create ZCXT Giftcheques  
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If


gstrTestCaseName = "Test_MD_AB_004 Create ZCXT Giftcheques"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_AB_004 Create ZCXT Giftcheques_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------XD01 -----------------------------------

Call SetComboByKey("RF02D-KTOKD",DT_XD01_7100_ACCOUNT_GROUP)
Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD01_7100_COMPANY_CODE,True)   
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD01_7100_SALES_ORGANIZATION,True) 
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD01_7100_DISTRIBUTION_CHANNEL,True)     
Call SetTextbox("Division","RF02D-SPART","",DT_XD01_7100_DIVISION,True) 
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)


Call SetComboByKey("SZA1_D0100-TITLE_MEDI",DT_XD01_0301_TITLE)
Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XD01_0301_NAME,False) 
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XD01_0301_SEARCH_TERM_12,False) 
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XD01_0301_STREETHOUSE_NUMBER,False) 
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XD01_0301_STREETHOUSE_NUMBER_OCC1,False) 
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XD01_0301_POSTAL_CODECITY,False) 
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XD01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XD01_0301_COUNTRY,False) 
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","",DT_XD01_0301_TELEPHONE,False) 

Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Control Data",False)
Call SetTextbox("VAT Reg\. No\.","KNA1-STCEG","",DT_XD01_7122_VAT_REG_NO,False) 
Call TakeScreenShot()
Call ClickButton("Company Code Data   \(Ctrl\+F2\)",False)

Call SetTextbox("Recon\. account","KNB1-AKONT","",DT_XD01_7211_RECON_ACCOUNT,False) 
Call SetTextbox("Sort key","KNB1-ZUAWA","",DT_XD01_7211_SORT_KEY,False)
Call SetTextbox("Cash mgmt group","KNB1-FDGRV","",DT_XD01_7211_CASH_MGMT_GROUP,False) 
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Payment Transactions",False)
Call  SelectCheckbox("KNB1-XZVER",0,DT_XD01_7215_PAYMENT_HISTORY_RECORD,False)
Call SetTextbox("Terms of payment","KNB1-ZTERM","",DT_XD01_7215_TERMS_OF_PAYMENT,False) 
Call SetTextbox("Payment methods","KNB1-ZWELS","",DT_XD01_7216_PAYMENT_METHODS,False) 
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Correspondence",False)
Call SetTextbox("Dunn\.Procedure","KNB5-MAHNA","",DT_XD01_7220_DUNNPROCEDURE,False) 
Call SetTextbox("Dunning Level","KNB5-MAHNS","",DT_XD01_7220_DUNNING_LEVEL,False) 
Call SetTextbox("Dunning clerk","KNB5-BUSAB","",DT_XD01_7220_DUNNING_CLERK,False)
Call SetTextbox("Clerk Abbrev\.","KNB1-BUSAB","",DT_XD01_7221_ACCTG_CLERK,False) 
Call SetTextbox("Account Statement","KNB1-XAUSZ","",DT_XD01_7221_ACCOUNT_STATEMENT,False)
Call PressEnter()
Call TakeScreenShot()


Call ClickButton("Sales Area Data   \(Ctrl\+F3\)",False)
Call SetTextbox("Cust\.pric\.proc\.","KNVV-KALKS","",DT_XD01_7311_CUSTPRICPROC,False) 
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Billing Documents",False)
Call SetTableData("SAPMF02DTCTRL_STEUERN","Tax classification","1","","",DT_XD01_7350_TABLECELL_TAX_CLASSIFICATION_0,False)
Call SetTextbox("Incoterms","KNVV-INCO1","",DT_XD01_7321_INCOTERMS,False) 
Call SetTextbox("Terms of payment","KNVV-ZTERM","",DT_XD01_7321_TERMS_OF_PAYMENT,False) 
Call SetTextbox("Acct assgmt group","KNVV-KTGRD","",DT_XD01_7322_ACCT_ASSGMT_GROUP,False) 
Call PressEnter()
Call TakeScreenShot()

Call SelectMenuBar("Extras;Classification")
Call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class","1","","",DT_XD01_1600_TABLECELL_CLASS_0,False)
Call PressEnter()
Call SetTableData("SAPLCTMSCHARS_S","Value","10","","",DT_XD01_5100_TABLECELL_VALUE_3,False)
''Call SetTableData("SAPLCTMSCHARS_S","Value","22","","",DT_XD01_5100_TABLECELL_VALUE_3_OCC1,False) 
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)

 Call ClickButton("Save   \(Ctrl\+S\)",False)
 Call GetStatusBar("item1","DT_CUSTOMER_OUTPUT")
Call VerifyStatusBar("Customer "& DT_CUSTOMER_OUTPUT &" has been created for company code GR02 sales area GS02 19 01")
 Call ClickButton("Cancel   \(F12\)",True)

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


