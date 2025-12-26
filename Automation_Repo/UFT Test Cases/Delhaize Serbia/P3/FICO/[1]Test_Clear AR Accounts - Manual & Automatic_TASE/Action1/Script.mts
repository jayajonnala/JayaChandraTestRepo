'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

'.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name       : Test_Clear AR Accounts - Manual & Automatic_TASE
'.................Author : TCS           : Bitan
'................ Creation Date          : 12th May
'.................Modified By            :
'.................Modified Date/Details  :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear AR Accounts - Manual & Automatic_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AR Accounts - Manual & Automatic_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


''''Login'''
'DataRowSet=2

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'''----------------------Tcode FB70----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB70_1000_COMPANY_CODE)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

''''Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
''''Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB70_1000_COMPANY_CODE)
'''''Capture the screenshot
''''Call TakeScreenShot()
''''Call ClickButtonIfExist("Continue   \(Enter\)",True)
''''
'''''Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
'''''Call PressEnter() 
'''''Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace((DT_FB70_0510_INVOICE_DATE),"/","."),False)
'''''Call PressEnter() 
'''''Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)
'''''Call PressEnter() 
'''''Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
'''''Call PressEnter() 
'''''Call SelectTab("TS","Basic data",False)
'''''Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)
''''''Select Calculate Tax field as true
'''''Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)
'''''Call PressEnter() 
'''''Call PressEnter() 

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDateFormat(DT_FB70_0510_INVOICE_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE,False)

Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT,False)

Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)
Call TakeScreenShot()
Call PressEnter() 
Call SelectTab("TS","Details",False)
Call PressEnter() 

Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
'Call PressEnter() 
Call PressEnter() 
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB70_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB70_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB70_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB70_0100_TABLECELL_COST_CENTER_0,False)
Call PressEnter() 
Call TakeScreenShot()

CAll SelectMenuBar("Document;Simulate")
Wait(2)
Call PressEnter()
Call TakeScreenShot()
Call VerifyWindowValue(DT_FB70_0120_CHECK_TEXT_OF_TITL)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_GETDOC_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_CHECKDOC)
Call TakeScreenShot()
Call ClickButtonIfExist("Cancel   \(F12\)",True)


'''----------------------Tcode FB75----------------------------

Call SetTcode(DT_FB70_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_FB70_0100_OKCD)
Call TakeScreenShot()

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB70_1000_COMPANY_CODE)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call ClickButtonIfExist("Switch Company Code   \(F7\)",False)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB70_1000_COMPANY_CODE)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER_OCC1,False)
Call PressEnter() 
Call SetTextbox("Document date","INVFO-BLDAT","",ConvertDateFormat(DT_FB70_0510_DOCUMENT_DATE),False)
Call PressEnter() 
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB70_0510_REFERENCE_OCC1,False)
Call PressEnter() 
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB70_0550_HEADERTEXT,False)
Call PressEnter() 
Call SelectTab("TS","Basic data",False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB70_0510_AMOUNT_OCC1,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX_OCC1,False)
Call PressEnter() 

Call PressEnter() 

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB70_0100_TABLECELL_GL_ACCT_0_OCC1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0_OCC2,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Tax code",1,DT_FB70_0100_TABLECELL_TAX_CODE_0_OCC1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Business area",1,DT_FB70_0100_TABLECELL_BUSINESS_AREA_0_OCC1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Cost center",1,DT_FB70_0100_TABLECELL_COST_CENTER_0_OCC1,False)
Call PressEnter() 
Call TakeScreenShot()

CAll SelectMenuBar("Document;Simulate")
Wait(2)
Call PressEnter()
Call TakeScreenShot()
Call VerifyWindowValue(DT_FB70_0120_CHECK_TEXT_OF_TITL_OCC1)


Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_GETDOC_1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_CHECKDOC_1)
Call TakeScreenShot()


Call LogOff()
Call FinalStatus ()

