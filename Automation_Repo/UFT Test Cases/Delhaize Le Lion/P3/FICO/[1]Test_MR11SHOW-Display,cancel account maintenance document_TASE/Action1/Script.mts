
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_MR11SHOW-Display,cancel acc doc"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
''''SAP Login'''
''
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

Call SetTextbox("Acct maint\. document","KBKP-BELNR","",DT_MR11SHOW_0100_ACCT_MAINT_DOCUMENT,False) 
Call SetTextbox("Fiscal year","KBKP-GJAHR","",DT_MR11SHOW_0100_FISCAL_YEAR,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickLabel("1", 0, False)

Call ClickButton("Reverse Account Maintenance Document   \(Shift\+F5\)",False)

Call SetTextBox("Posting Date","MR11_HEAD-BUDAT",0,ConvertDate(DT_MR11SHOW_0200_POSTING_DATE),False)

Call ClickButton("Reverse   \(Shift\+F1\)",False)

Call GetGridContent("", 0, "Message Text", 1, "Application Area", "M8", "DT_DOCUMENT_NR_OUTPUT")

Call ClickButton("Back   \(F3\)",False)

Call ClickButton("Back   \(F3\)",False)


'Call CheckExpectedTcode

Call SetTcode(DT_MR11SHOW_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_MR11SHOW_0003_PUR_ORDER,True)
Call TakeScreenShot()
Call PressEnter()

Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-EBELN", 0, DT_MR11SHOW_0003_PUR_ORDER, False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


