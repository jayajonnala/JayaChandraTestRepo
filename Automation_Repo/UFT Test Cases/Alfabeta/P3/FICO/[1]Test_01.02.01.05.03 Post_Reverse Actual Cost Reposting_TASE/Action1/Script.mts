

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.02.01.05.03 Post_Reverse Actual Cost Reposting_TASE
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_01.02.01.05.03 Post_Reverse Actual Cost Reposting_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''


''----------------------Tcode FB08----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

'''''''''
Call SetTextbox("Posting period","BSIS-MONAT","",DT_FB08_0030_GRIDCELL_5_PERIOD,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_FB08_0105_REVERSAL_REASON,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_FB08_0105_FISCAL_YEAR,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

Call FocusTextBox("Posting period","BSIS-MONAT",False)
Call ClickButton("Display document before reversal   \(F5\)",False)

Call VerifyTextBoxContent("Document Number","BKPF-BELNR",0,DT_FB08_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Yes",True)

'----------------------Tcode FAGLB03----------------------------

Call SetTcode(DT_FB08_0100_OKCD)
Call PressEnter() 

Call SetTextbox("Account Number","RACCT-LOW","",DT_FB08_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FB08_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FB08_1000_FISCAL_YEAR,False)
Call TakeScreenShot()
Call FocusTextBox("Fiscal Year","RYEAR",False)
Call ClickButton("Execute   \(F8\)",False)

Call GetGridContentByTitle("",0,"Debit",DT_INDEX,"DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT_OUTPUT",DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("",0,"Credit",DT_INDEX,"DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT_OUTPUT",DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetGridContentByTitle("",0,"Balance",DT_INDEX,"DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BALANCE_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BALANCE_OUTPUT",DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BALANCE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyGridCellContent("", DT_INDEX, "Debit", 0, DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DEBIT)
Call VerifyGridCellContent("",DT_INDEX, "Credit", 0, DT_FB08_0030_CHECK_GETCELLVALUE_OF_GRIDCELL_5_CREDIT)

Call SelectRowGuiGridbyRowNo("",0,DT_INDEX,False)
Call DoubleClickGuiGridCell("",0,DT_INDEX,"Period",False)
Call TakeScreenShot()

''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()


