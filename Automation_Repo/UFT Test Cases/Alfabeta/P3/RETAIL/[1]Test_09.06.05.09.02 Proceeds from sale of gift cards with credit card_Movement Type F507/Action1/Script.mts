

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

gstrTestCaseName = "Test_09.06.05.09.02 Proceeds from sale of gift cards with credit card_Movement Type F507"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.09.02 Proceeds from sale of gift cards with credit card_Movement Type F507_TASE.xls"


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

call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_FBL5N_1000_CUSTOMER_ACCOUNT,False)   
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_FBL5N_1000_COMPANY_CODE,False)   
''Call SetTextbox("to","SO_BUDAT-HIGH","",convertdate(DT_FBL5N_1000_POSTING_DATE),False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FBL5N_1000_POSTING_DATE),False)
wait(5)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",False)
call ClickButtonIfExist("Continue   \(Enter\)",true)

call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenSHot()
Call ClickButtonToolBar("&FIND", 0)
Call SetTextBox("Search Term:","GS_SEARCH-VALUE",0,DT_FBL5N_0841_SEARCH_TERM,True)
Call SelectCheckBox("GS_SEARCH-EXACT_WORD",0,"ON",True)
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenSHot()
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenSHot()
Call ClickButton("Transfer   \(Enter\)",True)

call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenSHot()
Call ClickButtonToolBar("&FIND", 0)
Call SetTextBox("Search Term:","GS_SEARCH-VALUE",0,DT_FBL5N_0841_SEARCH_TERM_OCC1,True)
Call SelectCheckBox("GS_SEARCH-EXACT_WORD",0,"ON",True)
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenSHot()
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenSHot()
Call ClickButton("Transfer   \(Enter\)",True)

Call  SelectColumnGuiGrid("", 1, "Text", False)
Call TakeScreenSHot()
call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call TakeScreenSHot()
Call SetTextboxNoName("Text","",DT_FBL5N_1105_TEXT,False)
call ClickButton("Execute   \(Enter\)",False)

'''Call  SelectColumnGuiGrid("", 1, "Reference", False)
'''Call TakeScreenSHot()
'''call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'''Call TakeScreenSHot()
'''Call SetTextboxNoName("Text","",DT_FBL5N_1105_REFERENCE,False)
'''call ClickButton("Execute   \(Enter\)",False)

call ClickButton("Display Document   \(Shift\+F2\)",False)
Call TakeScreenSHot()
call ClickButton("Call Up Document Overview   \(F9\)",False)

CAll GetTextboxValue("BKPF-XBLNR", "", "DT_FBL5N_1105_REFERENCE_OUTPUT", False)
Call TakeScreenSHot()
Call VerifyTextboxContent("Reference","BKPF-XBLNR","",DT_FBL5N_1105_REFERENCE_OUTPUT,False)

call VerifyGridCellContent("",1,"BSCHL",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"BSCHL",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

call VerifyGridCellContent("",1,"PRCTR",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
call VerifyGridCellContent("",2,"PRCTR",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

call VerifyGridCellContent("",1,"KTONR",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"KTONR",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

call VerifyGridCellContent("",1,"AZBET",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
call VerifyGridCellContent("",2,"AZBET",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)


call VerifyGridCellContent("",1,"SGTXT",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
call VerifyGridCellContent("",2,"SGTXT",0,DT_FBL5N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call LogOff()
Call FinalStatus ()

