'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AP Guarantees_p2_TASE 
'.................Author : TCS        :Bitan
'................ Creation Date    : 09th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage AP Guarantees_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AP Guarantees_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB08----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_FB08_0105_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Create Document List/Find Documents   \(F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_FB08_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_FB08_1000_FISCAL_YEAR,False)
Call SetTextbox("Ledger","BR_RLDNR-LOW","",DT_FB08_1000_LEDGER,False)
Call SetTextbox("Document type","BR_BLART-LOW","",DT_FB08_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Entry date","BR_CPUDT-LOW","",Replace((DT_FB08_1000_ENTRY_DATE),"/","."),False)
Call SetTextbox("Reference number","BR_XBLNR-LOW","",DT_FB08_1000_REFERENCE_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

''***** Document list section is not in editable mode, so edit is not possible, same is in tao log screen shot no data changed. *****
'Call SetGridData(gridTitle,1,"Document Header Text",DT_FB08_0500_GRIDCELL_0_DOCUMENT_HEADER_TEXT,False)
''Capture the screenshot
'Call TakeScreenShot()

Call DoubleClickGuiGridCell("","",1,"Document Header Text",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_FB08_0105_REVERSAL_REASON,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display document before reversal   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call PressEnter()
'Validate If doc is generated
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB08_0105_CHECK_TEXT_OF_STATUSBAR)

''display doc
Call SelectMenuBar("Document;Display")
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
Wait(1)
'validate grid components
call VerifyGridCellContent("",1,"Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",2,"Account","",DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


