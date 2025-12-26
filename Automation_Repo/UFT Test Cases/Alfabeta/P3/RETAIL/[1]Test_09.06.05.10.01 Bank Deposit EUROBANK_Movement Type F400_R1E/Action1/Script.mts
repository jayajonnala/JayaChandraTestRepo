

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

gstrTestCaseName = "Test_09.06.05.10.01 Bank Deposit EUROBANK_Movement Type F400_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.10.01 Bank Deposit EUROBANK_Movement Type F400_R1E_TASE.xls"


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

'//-----------------------------------/POSDW/MON0 -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
''''''''''''''''

call SelectRadioButton("X_AISEL","All items",false)

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
Call TakeScreenShot()
call ClickButton("Custom Selections   \(Ctrl\+F1\)",fALSE)
Call TakeScreenShot()
call ActivateNodeGuiTree(1,"General Ledger Line Items;Profit Center")
Call SetTextbox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
Call TakeScreenShot()
call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SelectColumnGuiGrid("", 1, "Text", False)
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenShot()
Call SetTextboxNoName("Text","",DT_FAGLL03_1105_TEXT,False)
Call TakeScreenShot()
call ClickButton("Execute   \(Enter\)",False)
Call TakeScreenShot()

''''Call SelectColumnGuiGrid("", 1, "Reference", False)
''''call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
''''Call TakeScreenShot()
''''Call SetTextboxNoName("Reference","",DT_FAGLL03_1105_REFERENCE,False)
''''Call TakeScreenShot()
''''call ClickButton("Execute   \(Enter\)",False)
''''Call TakeScreenShot()

call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Call TakeScreenShot()
call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"BSCHL",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL",1,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("",1,"KTONR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"KTONR",1,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("",1,"AZBET",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET",1,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("",1,"PRCTR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",2,"PRCTR",1,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call VerifyGridCellContent("",1,"SGTXT",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("",2,"SGTXT",1,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call VerifyGridCellContent("",1,"ZUONR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"ZUONR",1,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)

Call LogOff()
Call FinalStatus ()
