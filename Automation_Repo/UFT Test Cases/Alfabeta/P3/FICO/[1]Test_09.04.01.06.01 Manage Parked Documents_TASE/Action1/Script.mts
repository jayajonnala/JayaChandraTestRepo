

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.06.01 Manage Parked Documents
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

gstrTestCaseName = "Test_09.04.01.06.01 Manage Parked Documents"
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

'''--------TransactionCode-FB60 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB60_0010_REFERENCE",Cint(DT_FB60_0010_REFERENCE)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Vendor","INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB60_0010_REFERENCE, False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()

Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB60_0010_TEXT, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", "1", "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB60_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Park   \(Ctrl\+Shift\+F8\)",False)

Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
'Call VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was parked" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NUMBER_OUTPUT",DT_DOC_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)",False)
'
''''--------TransactionCode-FBV0 ----------''''

Call SetTcode(DT_FB60_1100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Doc\. Number","RF05V-BELNR","",DT_FB60_0100_DOC_NUMBER,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Shift\+F11\)",False)
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Information Message", "MESSTXT1", 0, Lcase(DT_FB60_0010_CHECK_TEXT_OF_MESSTXT1), False)
Call VerifyTextBoxContent("Information Message", "MESSTXT2", 0, Lcase(DT_FB60_0010_CHECK_TEXT_OF_MESSTXT2), False)

Call PressEnter()
Call TakeScreenShot

Call LogOff'
Call FinalStatus()
