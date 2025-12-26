
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p10
'.................Author : TCS        :Jaya
'................ Creation Date    : 
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Manual Customer Docs (Invoices, Credit, Debit Note)_p10_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''''--------TransactionCode-FB70----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_Company_Code,True)
'Call SetTextboxPopupIfExist("Company Code","BKPF-BUKRS","",DT_Company_Code)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Invoice date","INVFO-BLDAT", "", ConvertDate(DT_FB70_0510_INVOICE_DATE), False)
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB70_0510_CUSTOMER,False)
Call SetTextbox("Reference","INVFO-XBLNR", "", DT_FB70_0510_REFERENCE, False)
Call SetTextbox("Posting Date","INVFO-BUDAT", "", ConvertDate(DT_FB70_0510_POSTING_DATE), False)
Call SetTextbox("Amount","INVFO-WRBTR", "", DT_FB70_0510_AMOUNT, False)
Call SelectCheckBox("INVFO-XMWST",0,DT_FB70_0510_CALCULATE_TAX,False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetComboByKey("INVFO-BLART",DT_DOC_TYPE)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB70_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", "1", "", "", DT_FB70_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB70_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB70_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB70_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

'After Pressing Enter the Details screen is getting displayed automatically so commenting the below line.
'Call SelectTab("TS","Details",False)

Call SetTextbox("HeaderText","INVFO-BKTXT", "", DT_FB70_0550_HEADERTEXT, False)
Call TakeScreenShot
Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call SelectTab("TS","Tax",False)
Call TakeScreenShot
Call SelectTab("TS","Notes",False)
Call TakeScreenShot

Call SetTextArea(DT_FB70_0540_TEXTEDIT_SHELL)
Call TakeScreenShot

Call SelectTab("TS","Basic data",False)
Call TakeScreenShot

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code RS01")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB70_1200_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call ClickButtonIfExist("Cancel   \(F12\)",False)

Call LogOff'
Call FinalStatus()

