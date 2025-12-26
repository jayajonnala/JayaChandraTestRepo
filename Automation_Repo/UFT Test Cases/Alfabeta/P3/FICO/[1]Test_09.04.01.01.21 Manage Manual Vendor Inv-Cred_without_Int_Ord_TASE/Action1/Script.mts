'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.01.21 Manage Manual Vendor Inv-Cred_without_Int_Ord
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

gstrTestCaseName = "Test_09.04.01.01.21 Manage Manual Vendor Inv-Cred_without_Int_Ord"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''--------TransactionCode-FB65 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB65_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextbox("Document date", "INVFO-BLDAT", "", ConvertDate(DT_FB65_0010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date", "INVFO-BUDAT", "", ConvertDate(DT_FB65_0010_POSTING_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB65_0010_REFERENCE, False)
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB65_0010_TEXT, False)
Call TakeScreenShot


Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB65_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB65_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", "1", "", "", DT_FB65_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB65_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB65_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB65_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar(DT_FB65_1100_CHECK_TEXT_OF_STATUSBAR)

Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", "", False)
Call SetTableData("SAPLFSKBTABLE", "Order", "1", "", "", DT_FB65_0100_TABLECELL_ORDER_0, False)
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call SetTextbox("BaselineDt","INVFO-ZFBDT", "", ConvertDate(DT_FB65_0020_BASELINEDT), False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SelectTab("TS","Basic data",False)
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WRBTR", "", DT_FB65_0010_AMOUNT, False)
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Change Current Layout   \(Ctrl\+F8\)",False)
Call TakeScreenShot
Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR", "", DT_FB65_0850_FIND, True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)",True)
Call VerifyStatusBarMessageType("S")

Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_CNHX7832)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_NO_NAME_OCC3)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_N0)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
Call VerifyifGuiLabelExists(DT_FB65_0120_CHECK_TEXT_OF_NO_NAME_OCC5)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB65_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)",False)

Call LogOff'
Call FinalStatus()
