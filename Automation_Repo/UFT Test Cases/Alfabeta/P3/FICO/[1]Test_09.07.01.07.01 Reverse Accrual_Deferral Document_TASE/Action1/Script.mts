

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.07.01 Reverse Accrual_Deferral Document
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

gstrTestCaseName = "Test_09.07.01.07.01 Reverse Accrual_Deferral Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''''''----------------------Tcode FBS1----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 
Call TakeScreenShot()
 
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FBS1_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBS1_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FBS1_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBS1_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FBS1_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FBS1_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FBS1_0100_REFERENCE,False)
Call SetTextbox("Reversal Reason","BKPF-STGRD","",DT_FBS1_0100_REVERSAL_REASON,False)
Call SetTextbox("Reversal date","BKPF-STODT","",ConvertDate(DT_FBS1_0100_REVERSAL_DATE),False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FBS1_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FBS1_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FBS1_1007_BUSINESS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBS1_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_0300_ACCOUNT,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_0300_AMOUNT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBS1_0300_TEXT,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)

Call VerifyTextBoxContent("C", "RF05A-AZSAL",0,DT_FBS1_0700_CHECK_TEXT_OF_C, False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call PressEnter() 
Call TakeScreenShot()

Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FBS1_0100_CHECK_TEXT_OF_STATUSBAR)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FBS1_0100_POSTING_DATE_OCC1),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBS1_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FBS1_0100_TYPE_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBS1_0100_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_0100_ACCOUNT_OCC1,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FBS1_0100_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FBS1_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FBS1_0100_REFERENCE_OCC1,False)
Call SetTextbox("Reversal Reason","BKPF-STGRD","",DT_FBS1_0100_REVERSAL_REASON_OCC1,False)
Call SetTextbox("Reversal date","BKPF-STODT","",ConvertDate(DT_FBS1_0100_REVERSAL_DATE_OCC1),False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_0300_AMOUNT_OCC2,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FBS1_0300_TAX_CODE_OCC1,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FBS1_1007_COST_CENTER_OCC1,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FBS1_1007_BUSINESS_AREA_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FBS1_0300_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_0300_ACCOUNT_OCC1,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_0300_AMOUNT_OCC3,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBS1_0300_TEXT_OCC1,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call VerifyTextBoxContent("C", "RF05A-AZSAL",0,DT_FBS1_0700_CHECK_TEXT_OF_C_OCC1, False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call PressEnter() 
Call TakeScreenShot()

Call GetStatusBar("item1","DT_DOC_2_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_2_OUTPUT",DT_DOC_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FBS1_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

'''''''----------------------Tcode F.81----------------------------
Call SetTcode(DT_FBS1_0100_OKCD) 
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Company Code","BUKRS-LOW","",DT_FBS1_1000_COMPANY_CODE,False)

Call ClickButton("%_BELNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBS1_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBS1_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call SetTextbox("Fiscal Year","GJAHR-LOW","",Year(DT_FBS1_1000_FISCAL_YEAR),False)
Call SetTextbox("Document type","BLART-LOW","",DT_FBS1_1000_DOCUMENT_TYPE,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  
Call ClickButton("Reverse Documents   \(F6\)",False)
Call TakeScreenShot  

Call GetLabelContentByRefLabel("Company Code", -7, -64, "DT_MESSAGE_TEXT_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MESSAGE_TEXT_OUTPUT",DT_MESSAGE_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetLabelContentByRefLabel("Company Code", -7, -96, "DT_MESSAGE_TEXT_OCC1_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MESSAGE_TEXT_OCC1_OUTPUT",DT_MESSAGE_TEXT_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyifGuiLabelExistsByRelativeid((DT_FBS1_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_900014257),"wnd\[0\]/usr/lbl\[18,9\]")
Call VerifyifGuiLabelExistsByRelativeid((DT_FBS1_0120_CHECK_TEXT_OF_REVERSED_WITH_DOCUMENT_900014258),"wnd\[0\]/usr/lbl\[18,11\]")
'
''''''''--------TransactionCode-FAGLL03----------''''

Call SetTcode(DT_FBS1_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL", "All Items", False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FBS1_0100_POSTING_DATE),False)
Call ClickButton("%_SD_SAKNR_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBS1_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBS1_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#4;#1")

Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBS1_3010_TABLECELL_SINGLE_VALUE_0_OCC2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FBS1_3010_TABLECELL_SINGLE_VALUE_1_OCC2,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  

Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 7, "DMSHB", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_DMSHB)
Call VerifyGridCellContent("", 9, "DMSHB", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_DMSHB)

Call VerifyGridCellContent("", 4, "ICO_AUGP", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ICO_AUGP)
Call VerifyGridCellContent("", 8, "ICO_AUGP", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_ICO_AUGP)
Call LogOff'
Call finalstatus


