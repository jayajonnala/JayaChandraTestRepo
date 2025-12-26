

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Reverse AR Document_p2
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
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Reverse AR Document_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
'----------------------Tcode FB75----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter the Company Code
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_COMPANY_CODE) 'Rename 'Val' with 'strTextboxValue'
Call TakeScreenShot()

'Click on Continue
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(2)
'
''Click on Continue
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Wait(2)

'Enter the Details
Call SetTextbox("Customer","INVFO-ACCNT","",DT_FB75_0510_CUSTOMER,FALSE)
Call SetTextbox("Document date","INVFO-BLDAT","",ConvertDate(DT_FB75_0510_DOCUMENT_DATE),FALSE)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB75_0510_REFERENCE,FALSE)
''Click on Continue
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Wait(2)

''Click on Continue
'Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Wait(2)
''Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB75_0510_AMOUNT,FALSE)
Call SelectCheckbox("INVFO-XMWST",0,"ON",False)
Call PressEnter()


'Navigate to the Details Tab
Call SelectTab("TS","Details",False)
Wait(1)
Call PressEnter()

''''Call SetCombo("INVFO-BLART","Customer Invoice")

'Enter the Header Details
Call SetTextbox("HeaderText","INVFO-BKTXT","",DT_FB75_0550_HEADERTEXT,FALSE)
Call TakeScreenShot()


'Navigate to the Details Tab
Call SelectTab("TS","Basic data",False)
Wait(1)
Call PressEnter()


'Enter the details
Call SetTableData("SAPLFSKBTABLE","G/L Acct","1","","",DT_FB75_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableData("SAPLFSKBTABLE","Amount in doc.curr.","1","","",DT_FB75_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableData("SAPLFSKBTABLE","Tax Code","1","","",DT_FB75_0100_TABLECELL_TAX_CODE_0,False)
Call SetTableData("SAPLFSKBTABLE","Business Area","1","","",DT_FB75_0100_TABLECELL_BUSINESS_AREA_0,False)
Call SetTableData("SAPLFSKBTABLE","Cost Center","1","","",DT_FB75_0100_TABLECELL_COST_CENTER_0,False)
Call TakeScreenShot()


'Navigate to the Details Tab
Call SelectTab("TS","Details",False)
Wait(1)
Call TakeScreenShot()

'Navigate to the Payment Tab
Call SelectTab("TS","Payment",False)
Wait(1)
Call TakeScreenShot()


Call SetTextbox("Bline Date","INVFO-ZFBDT","",ConvertDate(DT_FB75_0520_BLINE_DATE),FALSE)
Call TakeScreenShot()


'Navigate to the Amount split Tab
Call SelectTab("TS","Amount split",False)
Wait(1)
Call TakeScreenShot()


'Navigate to the Notes Tab
Call SelectTab("TS","Notes",False)
Wait(1)
Call TakeScreenShot()

Call SetTextArea(DT_FB75_0540_TEXTEDIT_SHELL)
Call SelectTab("TS","Payment",False)
Wait(1)
Call SelectTab("TS","Basic data",False)
Call GetTextboxValue("RF05A-AZSAL",0,"DT_BAL_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_AMT,False)
Call PressEnter() 
Call SelectTab("TS","Payment",False)
Call PressEnter() 

'Click on Simulate Document Posting
Call ClickButton("Simulate Document Posting   \(F9\)",False)
Wait(2)


'Click Post
Call ClickButton("Post   \(Ctrl\+S\)",False) 
wait 5
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" was posted in company code RS01")

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


