		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_2_09.06.02.01.01 Clear AR Accounts (Manual and Automatic)
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

gstrTestCaseName = "Test_PRE_2_09.06.02.01.01 Clear AR Accounts (Manual and Automatic)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-FB75----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB75_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document date","INVFO-BLDAT", "", ConvertDate(DT_FB75_0510_DOCUMENT_DATE), False)
Call SetTextbox("Customer","INVFO-ACCNT", "", DT_FB75_0510_CUSTOMER, False)
Call TakeScreenShot
Call SetComboByKey("Document type", DT_FB75_0510_DOCUMENT_TYPE)
Call TakeScreenShot
Call SetTextbox("Text","INVFO-SGTXT", "", DT_FB75_0510_TEXT, False)
Call SetComboByKey("Tax Amount", DT_FB75_0510_TAX_AMOUNT)
Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST",0,"ON", Fasle)
Call TakeScreenShot
Call SetTextbox("Reference","INVFO-XBLNR", "", DT_FB75_0510_REFERENCE, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB75_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB75_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB75_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB75_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call SendKey("{F2}")
Call TakeScreenShot

Call SetTextbox("HeaderText","INVFO-BKTXT", "", DT_FB75_0550_HEADERTEXT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call SetTextbox("Bline Date","INVFO-ZFBDT", "", ConvertDate(DT_FB75_0520_BLINE_DATE), False)
Call pressEnter()
Call TakeScreenShot
Call SelectTab("TS","Basic data",False)
Call TakeScreenShot

Call GetTextboxValue("RF05A-AZSAL",0,"DT_FB75_1200_CHECK_TEXT_OF_BAL_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB75_1200_CHECK_TEXT_OF_BAL_OUTPUT",DT_FB75_1200_CHECK_TEXT_OF_BAL)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB75_0510_AMOUNT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",DT_FB75_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)",False)


Call LogOff()
Call FinalStatus ()

