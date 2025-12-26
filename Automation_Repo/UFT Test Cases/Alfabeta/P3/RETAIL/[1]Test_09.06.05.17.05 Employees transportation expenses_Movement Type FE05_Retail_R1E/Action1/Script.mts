

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

gstrTestCaseName = "Test_09.06.05.17.05 Employees transportation expenses_Movement Type FE05_Retail_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\09.06.05.17.05 Employees transportation expenses_Movement Type FE05_Retail_R1E_TASE.xls"


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

'//-----------------------------------/FBAGLL03-----------------------------------

call SelectRadioButton("X_AISEL","All items",false)

Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
call ClickButton("Custom Selections   \(Ctrl\+F1\)",fALSE)
call ActivateNodeGuiTree(1,"General Ledger Line Items;Profit Center")
Call SetTextbox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
call ClickButton("Save   \(Ctrl\+S\)",false)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
Call  SelectColumnGuiGrid("", 1, "Text", False)
Call TakeScreenSHot()
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenSHot()
Call SetTextboxNoName("Text","",DT_FAGLL03_1105_TEXT,False)
call ClickButton("Execute   \(Enter\)",False)

Call TakeScreenSHot()
call ClickButton("Display Document   \(Ctrl\+Shift\+F7\)",fALSE)
Call TakeScreenSHot()
call ClickButton("Call Up Document Overview   \(F9\)",fALSE)


Call VerifyGridCellContent("",1,"BSCHL",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"BSCHL",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("",1,"PRCTR",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("",2,"PRCTR",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

call VerifyGridCellContent("",1,"KTONR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"KTONR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",3,"KTONR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

call VerifyGridCellContent("",1,"AZBET",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
call VerifyGridCellContent("",2,"AZBET",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
call VerifyGridCellContent("",3,"AZBET",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

call VerifyGridCellContent("",1,"PRCTR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
call VerifyGridCellContent("",2,"PRCTR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
call VerifyGridCellContent("",3,"PRCTR",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)

call VerifyGridCellContent("",1,"SGTXT",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
call VerifyGridCellContent("",2,"SGTXT",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
call VerifyGridCellContent("",3,"SGTXT",0,DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_SGTXT)

Call LogOff()
Call FinalStatus ()

