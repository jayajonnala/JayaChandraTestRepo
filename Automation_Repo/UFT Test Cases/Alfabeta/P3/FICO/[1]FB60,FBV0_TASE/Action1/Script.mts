		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :FB60,FBV0
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

gstrTestCaseName = "FB60,FBV0"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''--------TransactionCode-FB60 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_REF",Cint(DT_REF)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetComboByKey("Transactn", DT_FB60_1100_TRANSACTN)

Call SetTextboxNoLabel("INVFO-ACCNT","",DT_ACCNT,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_INV_DATE), False)
'Call SetTextbox("Posting Date","INVFO-BUDAT", "", ConvertDate(DT_PSTNG_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_REF, False)
Call SetComboByKey("INVFO-BLART", DT_FB60_0010_DOCUMENT_TYPE)
Call PressEnter()
Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB60_0010_TEXT, False)
Call TakeScreenShot

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", "1", "", "", DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB60_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Simulate Document Posting   \(F9\)",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Park   \(Ctrl\+Shift\+F8\)",False)

Call GetStatusBar("item1","DT_OUTPUT_DOC_OUTPUT")
Call VerifyStatusBar("Document "&DT_OUTPUT_DOC_OUTPUT&" GR02 was parked" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_OUTPUT_DOC_OUTPUT",DT_OUTPUT_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButtonIfExist("Cancel   \(F12\)",False)

'''--------TransactionCode-/nFBV0 ----------''''
Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Doc\. Number", "RF05V-BELNR", "", DT_FB60_0100_DOC_NUMBER, False)
Call SetTextbox("Company Code","RF05V-BUKRS", "", DT_FB60_0100_COMPANY_CODE, False)
Call SetTextbox("Fiscal Year","RF05V-GJAHR", "", Year(DT_FB60_0100_FISCAL_YEAR), False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Post   \(Shift\+F11\)",False)
Call PressEnter() 
Call TakeScreenShot
Call VerifyTextBoxContent("Information Message", "MESSTXT1", 0, Lcase(DT_FB60_0010_CHECK_TEXT_OF_MESSTXT1), True)

Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call LogOff
Call FinalStatus()


