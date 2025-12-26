

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_09.06.05.16.04 Card + Payment Deficit_Movement Type F408_Retail_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.16.04 Card + Payment Deficit_Movement Type F408_Retail_R1E_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------/FBL5N -----------------------------------

call SelectRadioButton("X_AISEL","All items",false)
Call TakeScreenSHot()
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
Call TakeScreenSHot()
call ClickButton("Custom Selections   \(Ctrl\+F1\)",fALSE)
Call TakeScreenSHot()
Call ActivateNodeGuiTree(1,"General Ledger Line Items;Profit Center")
Call SetTextbox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
Call ClickButton("Save   \(Ctrl\+S\)",false)
Call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
Call  SelectColumnGuiGrid("", 1, "Text", False)
Call TakeScreenSHot()
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenSHot()
Call SetTextboxNoName("Text","",DT_FAGLL03_1105_TEXT,False)
Call ClickButton("Execute   \(Enter\)",False)

Call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",fALSE)
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT", "", DT_FAGLL03_0300_CHECK_TEXT_OF_GL_ACCOUNT, False)
Call ClickButton("Call Up Document Overview   \(F9\)",fALSE)

Call VerifyGridCellContent("",1,"BSCHL","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("",1,"KTONR","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"KTONR","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("",1,"AZBET","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("",1,"PRCTR","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",2,"PRCTR","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

Call VerifyGridCellContent("",1,"SGTXT","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("",2,"SGTXT","",DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call LogOff()
Call FinalStatus ()
