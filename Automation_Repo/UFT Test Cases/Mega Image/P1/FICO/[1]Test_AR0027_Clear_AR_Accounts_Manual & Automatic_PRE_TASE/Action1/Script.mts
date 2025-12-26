'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : DT_AR0027_Clear_AR_Accounts_Manual & Automatic_anual_standard_TASE
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


gstrTestCaseName = "DT_AR0027"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''----------------------Tcode F-32----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'Call SelectRadioButton("RF05A-XPOS1","Document Number",False)
'Call SelectCheckbox("RF05A-XNOPS",0,DT_F22_131_NORMAL_OI,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F32_131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F32_131_COMPANY_CODE,False)
Call SetTextbox("Clearing Date","BKPF-BUDAT","",ConvertDate(DT_F32_131_CLEARING_DATE),False)

'Call TakeScreenShot()
'Call PressEnter()
Call TakeScreenShot()
'Click Process Open Items
Call ClickButton("Process Open Items   \(Shift\+F4\)",False) 
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Select All",False)
'''''Call ClickButton("Activate Items",False)
Call ClickButton("Deactivate Items",False)
Call TakeScreenShot()
Call ClickButton("Field content search",False)
Call TakeScreenShot()
Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call SetTextbox("From","RF05A-SEL01","0",DT_F32_0731_FROM,False)
Call SetTextbox("From","RF05A-SEL01","1",DT_F32_0731_FROM_OCC2,False)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call ClickButton("Select All",False)
Call ClickButton("Activate Items",False)
Call TakeScreenShot()
Call ClickButton("Document Overview   \(Shift\+F2\)",False)
Call SelectMenuBar("Document;Simulate")

Call ClickButton("Post   \(Ctrl\+S\)",False)
'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_F32_131_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Exit   \(Shift\+F3\)",False)
wait(2)
''----------------------Tcode FB03----------------------------

'Enter the Tcode
Call SetTcode(DT_F32_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F32_100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F32_100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F32_100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call DoubleClickGuiGridCell("",0, 1, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_F32_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call DoubleClickGuiGridCell("",0, 2, "Posting Key", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_F32_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
Call VerifyGridCellContent("",1,"Posting Key","",DT_F32_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","",DT_F32_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)

Call VerifyTextBoxContent("Document type","BKPF-BLART",0,DT_F32_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call TakeScreenShot()
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Call ClickButton("Exit   \(Shift\+F3\)",False)
wait(2)

'''----------------------Tcode FBL5N----------------------------
'Enter the Tcode
Call SetTcode(DT_F32_100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_F32_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_F32_1000_COMPANY_CODE,False)
Call SetTextbox("Clearing date","SO_AUGDT-LOW","",ConvertDate(DT_F32_1000_CLEARING_DATE),False)
Call SetTextbox("to","SO_AUGDT-HIGH","",ConvertDate(DT_F32_1000_TO_CL_DATE),False)

'Call SelectCheckbox("X_SHBV",0,"ON",False)
'Call SelectCheckbox("X_NORM",0,"OFF",False)
Call SelectRadioButton("X_CLSEL","Cleared items",False)

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
''''''Call SelectRowGuiTableByRow("SAPLSKBHTC_FIELD_LIST_820",11, True)
Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","",DT_F32_850_FIND_OCC2,True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel. fields \(CTRL\+F3\)", True)
Call ClickButton("Copy   \(Enter\)", True)
Call SetTextbox("Clearing Document","%%DYN001-LOW","",DT_F32_1105_CLEARING_DOCUMENT,True)
Call TakeScreenShot()
Call ClickButton("Execute   \(Enter\)", True)

'Call VerifyStatusBarMessageType("S")

Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC3,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC9,0)
'Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"/","."),0)
'Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"/","."),0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_DZ,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_CL,0)
'Call VerifyifGuiLabelExists_ByIndex("S_LEDR",0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_DR,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC4,0)
'Call VerifyifGuiLabelExists_ByIndex("S_LEDG",0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC5,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC6,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC7,0)
Call VerifyifGuiLabelExists_ByIndex(DT_F32_120_CHECK_TEXT_OF_NO_NAME_OCC8,0)
'Call VerifyifGuiLabelExists_ByIndex(Replace(DT_F22_0120_CHECK_TEXT_OF_NO_NAME_OCC7,"/","."),0)

'Call ClickButton("Back   \(F3\)",False)
'wait(2)
'
'Call ClickButton("Back   \(F3\)",False)
'wait(2)
'
'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

