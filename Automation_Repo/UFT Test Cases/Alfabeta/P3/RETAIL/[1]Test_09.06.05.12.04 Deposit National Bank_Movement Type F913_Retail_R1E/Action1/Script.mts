

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

gstrTestCaseName = "Test_09.06.05.12.04 Deposit National Bank_Movement Type F913_Retail_R1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.12.04 Deposit National Bank_Movement Type F913_Retail_R1E_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''------------Login-----------
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''------------Tcode----------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenSHot()
'''''//-----------------------------------/FBL5N -----------------------------------

call SelectRadioButton("X_AISEL","All items",false)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)   
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)   
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FBL5N_1000_POSTING_DATE),False)
wait(5)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
call ClickButtonIfExist("Continue   \(Enter\)",true)
call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",fALSE)
Call TakeScreenSHot()
call ClickButtonToolBar("&FIND",0)
Call TakeScreenSHot()
call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Reference",true)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD", 0, "ON", True)
TakeScreenShot()
call ClickButton("OK   \(Enter\)",true)
TakeScreenShot()
call ClickButton("Cancel   \(F12\)",true)
TakeScreenShot()
call ClickButton("Show Selected Fields \(F7\)",true)
call ClickButtonIfExist("Transfer   \(Enter\)",true)

Call SelectColumnGuiGrid("", 0, "Text", False)
Call TakeScreenSHot()
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",fALSE)
Call SetTextboxNoName("Text","",DT_FBL5N_1105_TEXT,False)
call ClickButton("Execute   \(Enter\)",fALSE)
Call TakeScreenSHot()

Call SelectColumnGuiGrid("", 0, "Reference", False)
Call TakeScreenSHot()
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",fALSE)
Call SetTextboxNoName("Reference","",DT_FBL5N_1105_REFERENCE,False)
call ClickButton("Execute   \(Enter\)",fALSE)
Call TakeScreenSHot()

call ClickButton("Display Document   \(Shift\+F2\)",fALSE)
Call TakeScreenSHot()
Call VerifyTextBoxContent("Customer", "KNA1-KUNNR", "", DT_FBL5N_0301_CHECK_TEXT_OF_CUSTOMER, False)
call ClickButton("Call Up Document Overview   \(F9\)",fALSE)
Call TakeScreenSHot()

call VerifyGridCellContent("",1,"BSCHL","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"BSCHL","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

call VerifyGridCellContent("",1,"PRCTR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
call VerifyGridCellContent("",2,"PRCTR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

call VerifyGridCellContent("",1,"AZBET","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
call VerifyGridCellContent("",2,"AZBET","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

call VerifyGridCellContent("",1,"KTONR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"KTONR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

call VerifyGridCellContent("",1,"SGTXT","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
call VerifyGridCellContent("",2,"SGTXT","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

call VerifyGridCellContent("",1,"ZUONR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
call VerifyGridCellContent("",2,"ZUONR","",DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)

Call LogOff()
Call FinalStatus ()

