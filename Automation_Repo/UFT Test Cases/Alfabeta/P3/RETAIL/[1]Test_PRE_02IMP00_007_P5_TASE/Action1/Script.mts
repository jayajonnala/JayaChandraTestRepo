	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_02IMP00_007_P5
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

gstrTestCaseName = "Test_PRE_02IMP00_007_P5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_02IMP00_007_P5.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

'//-----------------------------------MIGO -----------------------------------

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
call SetComboByKey("Transaction",DT_MIRO_6000_TRANSACTION)

Call ClickButton("btn\[12\]",fALSE)
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_6211_RM08MEBELN,False)  
Call TakeScreenShot
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call TakeScreenShot
Call SetTextboxNoLabel("RM08M-EBELN","",DT_MIRO_6211_RM08MEBELN,False)
Call SetComboByKey("RM08M-XWARE_BNK",DT_MIRO_6211_RM08MXWARE_BNK)
call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPLMR1MTC_MSEL_VEN", 1, True)
Call TakeScreenShot()
Call ClickButton("Continue   \(F8\)",False)
Call TakeScreenShot()
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",1,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",2,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",3,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",4,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",5,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call TakeScreenShot

Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call TakeScreenShot

Call WriteRunTimeDataToExcel("DT_GET_AMOUNT_OUTPUT", DT_GET_AMOUNT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)

Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call TakeScreenShot

Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",1,DT_MIRO_0010_CALCULATE_TAX,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",2,DT_MIRO_0010_CALCULATE_TAX,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",3,DT_MIRO_0010_CALCULATE_TAX,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",4,DT_MIRO_0010_CALCULATE_TAX,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",5,DT_MIRO_0010_CALCULATE_TAX,False)
Call TakeScreenShot

Call PressEnter()
Call TakeScreenshot()

'Call VerifyTextBoxContentIconName("RM08M-AMPEL", "", DT_MIRO_6000_CHECK_ICONNAME_OF_TRANSACTION, False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_Document_no_Output")
Call WriteRunTimeDataToExcel("DT_GET_AMOUNT_OUTPUT", DT_GET_AMOUNT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)
call VerifyStatusBar(LCase(DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR))
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()

