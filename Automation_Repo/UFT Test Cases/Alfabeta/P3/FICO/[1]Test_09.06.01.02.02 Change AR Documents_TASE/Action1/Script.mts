		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.01.02.02 Change AR Documents
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

gstrTestCaseName = "Test_09.06.01.02.02 Change AR Documents"
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

'''''--------TransactionCode-FB09----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB09_0102_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",Year(DT_FB09_0102_FISCAL_YEAR),False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB09_0102_COMPANY_CODE,False)
''text box label changed
''Call SetTextbox("Line item","RF05L-BUZEI","",DT_FB09_0102_LINE_ITEM,False)
Call SetTextboxNoLabel("RF05L-BUZEI","",DT_FB09_0102_LINE_ITEM,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Text","BSEG-SGTXT","",DT_FB09_0301_TEXT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_FB09_0102_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

'''''''--------TransactionCode-FBL5N----------''''

Call SetTcode(DT_FB09_0102_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FB09_1000_CUSTOMER_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN011_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB09_1106_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
'
Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenShot

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB09_1106_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB09_1106_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB09_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call TakeScreenShot

Call LogOff'
Call FinalStatus()
