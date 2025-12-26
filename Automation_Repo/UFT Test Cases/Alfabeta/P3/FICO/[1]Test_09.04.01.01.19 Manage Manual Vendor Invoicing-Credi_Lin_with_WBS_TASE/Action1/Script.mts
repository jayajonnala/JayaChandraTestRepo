

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.01.19 Manage Manual Vendor Invoicing-Credi_Lin_with_WBS
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

gstrTestCaseName = "Test_09.04.01.01.19 Manage Manual Vendor Invoicing-Credi_Lin_with_WBS"
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
'
''''--------TransactionCode-FB65 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextbox("Document date", "INVFO-BLDAT", "", ConvertDate(DT_FB65_0010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date", "INVFO-BUDAT", "", ConvertDate(DT_FB65_0010_POSTING_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB65_0010_REFERENCE, False)
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB65_0010_TEXT, False)
Call SetTextbox("Amount", "INVFO-WRBTR", "", DT_FB65_0010_AMOUNT, False)
Call TakeScreenShot


Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB65_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", "1", "", "", DT_FB65_0100_TABLECELL_DC_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", "1", "", "", DT_FB65_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB65_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB65_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB65_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "WBS element", "1", "", "", DT_FB65_0100_TABLECELL_WBS_ELEMENT_0, False)

Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call SetTextbox("BaselineDt", "INVFO-ZFBDT", "", ConvertDate(DT_FB65_0020_BASELINEDT), False)
Call PressEnter()
Call TakeScreenShot

Call SelectTab("TS","Details",False)
Call TakeScreenShot
Call SetTextbox("G/L", "INVFO-HKONT", "", DT_FB65_0050_GL, False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call PressEnter()
Call ClickButton("Change Current Layout   \(Ctrl\+F8\)",False)
Call ClickButton("B_SEARCH",True)

Call SetTextbox("Find","GD_SEARCHSTR","",DT_FB65_0850_FIND,True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot

Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_G2999955GZ000025OPEX)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButton("Cancel   \(F12\)",False)

Call LogOff'
Call Finalstatus()
