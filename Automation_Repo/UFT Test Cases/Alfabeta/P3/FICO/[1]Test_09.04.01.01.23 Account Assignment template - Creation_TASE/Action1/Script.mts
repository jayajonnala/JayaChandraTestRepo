

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.01.23 Account Assignment template - Creation
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


gstrTestCaseName = "Test_09.04.01.01.23 Account Assignment template - Creation"
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
''''--------TransactionCode-FB50 ----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INC_VAL",Cint(DT_INC_VAL)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

'Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB65_0010_VENDOR,False)
Call SetTextbox("Document date", "ACGL_HEAD-BLDAT", "", ConvertDate(DT_FB50_1010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date", "ACGL_HEAD-BUDAT", "", ConvertDate(DT_FB50_1010_POSTING_DATE), False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Reference", "ACGL_HEAD-XBLNR", "", DT_FB50_1010_REFERENCE, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB50_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", "1", "", "", DT_FB50_0100_TABLECELL_DC_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "2", "", "", DT_FB50_0100_TABLECELL_GL_ACCT_1, False)
Call SetTableData("SAPLFSKBTABLE", "D/C", "2", "", "", DT_FB50_0100_TABLECELL_DC_1, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "2", "", "", DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1, False)

Call SetTextbox("Document date", "ACGL_HEAD-BLDAT", "", ConvertDate(DT_FB50_1010_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date", "ACGL_HEAD-BUDAT", "", ConvertDate(DT_FB50_1010_POSTING_DATE), False)

Call SelectMenuBar("Edit;Acct assignment templates;Save Account Assignment Template")
Call SetTextbox("Act asst temp\.", "KOMU-KMNAM", "", DT_FB50_0210_ACT_ASST_TEMP,True)

Call ClickButton("Continue   \(Enter\)",True)
Call ActivateNodeGuiTree(0,"#2;#1")
Call ClickButton("Cancel",True)


'-------------------------FB50-----------------------------------------
Call SetTcode(DT_FB50_1001_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTcode(DT_FB50_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE_OCC1,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Display Templates   \(Shift\+F1\)",False)
Call ActivateNodeGuiTree(0,"#2;#1")

Call VerifyTableCellContent(1, "G/L acct", "SAPLFSKBTABLE", DT_FB50_0100_CHECK_TEXT_OF_TABLECELL_GL_ACCT_0)
Call VerifyTableCellContent(2, "G/L acct", "SAPLFSKBTABLE", DT_FB50_0100_CHECK_TEXT_OF_TABLECELL_GL_ACCT_1)
Call VerifyTableCellContent(1, "Amount in doc.curr.", "SAPLFSKBTABLE", DT_FB50_0100_CHECK_TEXT_OF_TABLECELL_AMOUNT_IN_DOCCURR_0)
Call VerifyTableCellContent(2, "Amount in doc.curr.", "SAPLFSKBTABLE", DT_FB50_0100_CHECK_TEXT_OF_TABLECELL_AMOUNT_IN_DOCCURR_1)

Call LogOff
Call FinalStatus()
