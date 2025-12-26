'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_FB08-Reverse manual posting and Non- PO invoices  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_FB08-Reverse post& Non- PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-FB08----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
'Call TakeScreenShot
'Call PressEnter()

Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)      
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BSIS-BUDAT","",ConvertDate(DT_FB08_0105_POSTING_DATE),False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_FB08_0105_REVERSAL_REASON,False)
Call SetTextbox("Posting period","BSIS-MONAT","",DT_FB08_0105_POSTING_PERIOD,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_FB08_0105_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButtonIfExist("Display document before reversal   \(F5\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1) 
Call VerifyGridCellContent("", 1, "Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "Account","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1", "DT_OP_DOC_NO")
'datatable_row = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_FB08_0105_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
''
''''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_FB08_0105_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Account Number","RACCT-LOW","",DT_FB08_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FB08_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FB08_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
''Need to write generic based on value presetn in Rows
Call DoubleClickGuiGridCell("","",DT_ROW,"Period",False)
Call TakeScreenShot

Call SelectColumnGuiGrid("", 0, "Document Number", False)
Call TakeScreenShot
Call SelectMenuBar("Settings;Switch List")
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB08_1105_DOCUMENT_NUMBER,True)
'Call TakeScreenShot
'Call ClickButton("Execute   \(Enter\)",True)
'Call TakeScreenShot
'Call VerifyGridCellContent("",1,"Clearing Document",0,DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FB08_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call VerifyifGuiLabelExists_ByIndex("S_LEDG", 0)
Call ClickButtonifexist("Last Column   \(Ctrl\+F4\)",False)
Call VerifyifGuiLabelExists_ByIndex(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL, 0)
Call TakeScreenShot
'
'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

