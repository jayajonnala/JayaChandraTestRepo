

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.01.05.01 Change (Non-PO Invoices)
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

gstrTestCaseName = "Test_09.04.01.05.01 Change (Non-PO Invoices)"
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

''''--------TransactionCode-FB50 ----------''''
'

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FBL1N_1000_VENDOR_ACCOUNT,False)
Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FBL1N_1000_COMPANY_CODE,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Wait(5)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)


Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBL1N_3010_TABLECELL_SINGLE_VALUE_0,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call ClickButton("Change document   \(Ctrl\+Shift\+F8\)",False)
Call TakeScreenShot

Call SetTextbox("Assignment","BSEG-ZUONR","",DT_FBL1N_0302_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBL1N_0302_TEXT,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call VerifyGridCellContent("", 1, "ZUONR", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER_OCC1)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FBL1N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)

Call Logoff'
Call FinalStatus()

