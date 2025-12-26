

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.01.02 Clear AP Accounts (Manual and Automatic)_P2
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

gstrTestCaseName = "Test_09.04.03.01.02 Clear AP Accounts (Manual and Automatic)_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.01 Manage Manual Post  Direct Domestic Vendor Invoic.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''''--------TransactionCode-FB60----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetComboByKey("Transactn",DT_FB60_1100_TRANSACTN)
Call TakeScreenShot

Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB60_0010_REFERENCE, False)
Call TakeScreenShot

Call SetComboByKey("INVFO-BLART",DT_FB60_0010_DOCUMENT_TYPE)
Call TakeScreenShot

Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot

Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB60_0010_TEXT, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax Code", "1", "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call PressEnter()

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call GetStatusBar("item1","DT_DOC_2_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_2_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_2_OUTPUT",DT_DOC_2_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'Call ClickButtonIfExist("Cancel   \(F12\)",False)

Call LogOff'
Call FinalStatus()

