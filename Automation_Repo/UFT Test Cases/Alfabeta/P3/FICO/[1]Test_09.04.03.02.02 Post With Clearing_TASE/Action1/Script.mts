

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.02.02 Post With Clearing
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

gstrTestCaseName = "Test_09.04.03.02.02 Post With Clearing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.03.02.02 Post With Clearing.xls"
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

''''''--------TransactionCode-FB60----------''''

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
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB60_0010_REFERENCE, False)
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB60_0010_TEXT, False)

Call SetComboByKey("Document type",DT_FB60_0010_DOCUMENT_TYPE)
Call TakeScreenShot
Call PressEnter


Call SetComboByKey("Document type",DT_FB60_0010_DOCUMENT_TYPE_OCC1)
Call TakeScreenShot
Call PressEnter

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", Replace(DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,".",","), False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", "1", "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB60_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call PressEnter
Call TakeScreenShot

Call SelectTab("TS","Details",False)
Call TakeScreenShot

Call SetTextbox("G/L", "INVFO-HKONT", "", DT_FB60_0050_GL, False)
Call TakeScreenShot

Call SelectTab("TS","Basic data",False)
Call TakeScreenShot

' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot

'If VerifyWindowValue("Enter Vendor Invoice: Company Code GR02") Then
	Call PressEnter() 
	Call PressEnter() 
	Call PressEnter() 
'End If

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_FB60_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB60_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call ClickButtonIfExist("Cancel   \(F12\)",True)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB60_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB60_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FBL1N----------''''

Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_0122_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB60_0122_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB60_0122_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB60_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB60_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB60_0122_DOCHEADER_TEXT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB60_0122_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB60_0122_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB60_0300_AMOUNT,False)
Call ClickButtonIfExist("Display Additional Data for Document Item   \(F7\)",False)
Call SetTextbox("House Bank","BSEG-HBKID","",DT_FB60_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_FB60_0330_BSEGHKTID,False)
Call TakeScreenShot

Call SelectMenuBar("Document;Simulate")
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_FB60_0710_ACCOUNT_TYPE,False)
Call SetTextbox("Company Code","RF05A-AGBUK","",DT_FB60_0710_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_FB60_0710_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter() 

Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Deactivate Items",False)
Call ClickButtonIfExist("Field content search",False)
Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call PressEnter() 
Call SetTextbox("From","RF05A-SEL01","",DT_FB60_0731_FROM,False)
Call PressEnter()
Call ClickButtonIfExist("Select All",False)
Call ClickButtonIfExist("Activate Items",False)
Call ClickButtonIfExist("Clearing Text\.\.\.   \(Ctrl\+F2\)",False)
Call PressEnter()
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB60_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB60_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB60_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB60_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FBL1N----------''''
Call SetTcode(DT_FB60_0122_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Layout","PA_VARI","",DT_LAYOUT,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB60_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot
Call SelectRadioButton("X_CLSEL","Cleared items",False)
Call TakeScreenShot
Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB60_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB60_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 3, "ICO_AUGP", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ICO_AUGP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "Document type", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "Document type", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 1, "Document Date", 0, ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 2, "Document Date", 0, ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 1, "Amount in local currency", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "Amount in local currency", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 1, "Local Currency", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "Local Currency", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
Call VerifyGridCellContent("", 1, "Clearing Document", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 2, "Text", 1, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call LogOff
Call FinalStatus()











