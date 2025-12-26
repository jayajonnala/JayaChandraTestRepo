

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.17.1. Copy Info Record from Vendor to Vendor
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.3.1.2.2. Modify Customer with Account Group ZCXT - DG External_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.3.1.2.2. Modify Customer with Account Group ZCXT - DG External_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_2.3.1.2.2. Modify Customer with Account Group ZCXT - DG External_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'''----------------------Tcode XD02----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextboxNoLabel("RF02D-KUNNR","",DT_XD02_7101_RF02DKUNNR,True)
Call SetTextbox("Company code","RF02D-BUKRS","",DT_XD02_7101_COMPANY_CODE,True)
Call SetTextbox("Sales Organization","RF02D-VKORG","",DT_XD02_7101_SALES_ORGANIZATION,True)
Call SetTextbox("Distribution Channel","RF02D-VTWEG","",DT_XD02_7101_DISTRIBUTION_CHANNEL,True)
Call SetTextbox("Division","RF02D-SPART","",DT_XD02_7101_DIVISION,True)
Call TakeScreenShot()

Call ClickButtonIfExist("Customer's sales areas\.\.\.",True)
Call TakeScreenShot()
Call SelectCellGuiTable("SAPMF02DTCTRL_KUNDENVERTRIEB","Distr. Chl","Distr. Chl","07",True)
'Call DoubleClick()
Call SendKey("{F2}")
wait 2
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XD02_0301_NAME,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XD02_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Mobile Phone","SZA1_D0100-MOB_NUMBER","",DT_XD02_0301_MOBILE_PHONE,False)
Call SetComboByKey("ADDR1_DATA-LANGU",DT_LANGUAGE)
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","",DT_XD02_0301_TELEPHONE,False)
Call TakeScreenShot()

Call FocusTextBox("Mobile Phone","SZA1_D0100-MOB_NUMBER",False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Payment Transactions",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLUNGSVERKEHR","Bank Account",1,DT_XD02_7131_TABLECELL_BANK_ACCOUNT_0,False)
Call TakeScreenShot()

Call ClickButton("Company Code Data   \(Ctrl\+F2\)",False)
Wait(2)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Payment Transactions",False)
Call TakeScreenShot()

Call FocusTextBox("Payment methods","KNB1-ZWELS",False)
Call SendKey("{F4}")
Wait(1)
Call SetTableDataNoRef("SAPMF02DTCTRL_ZAHLWEGE","Indicator: Payment mtd selected(outgoing)?",2,DT_XD02_1215_TABLECELL_INDICATOR_PAYMENT_MTD_SELECTED_OUTGOING__0,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Sales Area Data   \(Ctrl\+F3\)",False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Shipping",False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Billing Documents",False)

Call FocusTextBox("Terms of payment","KNVV-ZTERM",False)
Call SendKey("{F4}")
Wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call ClickLabel(DT_XD02_LABEL_1,"",True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(3)
Call VerifyStatusBar(DT_EXPECTED)

Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call LogOff()
Call FinalStatus ()
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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

gstrTestCaseName = "Test_2.1.1.17.1. Copy Info Record from Vendor to Vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode ZMDPU_INFOREC_COPY----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call CheckScreen(DT_SAPTRANSACTIONCODE,Lcase(DT_ZMDPU_INFOREC_COPY_1000_CHECK_TEXT_OF_TITL))

'''Call SelectCheckbox("P_NOTDEL",0,DT_ZMDPU_INFOREC_COPY_1000_DO_NOT_COPY_INFORECORDS_MARKED_FOR_DELETION_,False)
Call SelectCheckbox("P_MRKORI",0,DT_ZMDPU_INFOREC_COPY_1000_MARK_ORIGINAL_INFORECORD_AS_DELETED_,False)
Call SelectCheckbox("P_COPINF",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_INFORECORD_AND_PO_TEXT_,False)
Call SelectCheckbox("P_COPSUB",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_SUBRANGE_ASSIGNMENT_,False)
Call SelectCheckbox("P_MRKGTI",0,DT_ZMDPU_INFOREC_COPY_1000_MARK_GTINS_AS_TARGET_VENDOR_SPECIFIC_,False)
Call SelectCheckbox("P_COPPRI",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_PRICING_CONDITIONS,False)
Call SelectCheckbox("P_MRKVEN",0,DT_ZMDPU_INFOREC_COPY_1000_MARK_NEW_VENDOR_AS_REGULAR_,False)
Call SelectCheckbox("P_SRCVEN",0,DT_ZMDPU_INFOREC_COPY_1000_USE_SOURCE_VENDOR_AS_PRIOR_VENDOR_,False)
Call SelectCheckbox("P_COPVEN",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_VENDOR_ARTICLE_CODE_,False)

Call SetTextbox("Purch\. Organization","P_EKORG","",DT_ZMDPU_INFOREC_COPY_1000_PURCH_ORGANIZATION,False)
Call SetTextbox("Source Vendor","P_SRC_V","",DT_ZMDPU_INFOREC_COPY_1000_SOURCE_VENDOR,False)
Call SetTextbox("Target Vendor","P_TGT_V","",DT_ZMDPU_INFOREC_COPY_1000_TARGET_VENDOR,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDPU_INFOREC_COPY_1000_ARTICLE,False)

Call TakeScreenShot()

Call PressEnter() 

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call SelectRowGuiGrid("",0,"Purchasing Organization",DT_ZMDPU_INFOREC_COPY_1000_PURCH_ORGANIZATION,False)

Call TakeScreenShot()

Call ClickButton("PROCESS   \(F9\)",False)

Call TakeScreenShot()

Call ClickButton("YES",True)

Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_ZMDPU_INFOREC_COPY_0500_CHECK_TEXT_OF_STATUSBAR)

Call SelectRowGuiGridbyRowNo("",0,1,False)
Call DoubleClickGuiGridCell("",0,1,"Inforecord number",False)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call VerifyTextBoxContent("Info record","EINA-INFNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_INFO_RECORD,False)
Call VerifyTextBoxNoLabelContent("EINA-LIFNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_VENDOR,False)
'Call VerifyTextBoxContent("Vendor","EINA-LIFNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_VENDOR,False)
Call VerifyTextBoxContent("Article","EINA-MATNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_ARTICLE,False)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
