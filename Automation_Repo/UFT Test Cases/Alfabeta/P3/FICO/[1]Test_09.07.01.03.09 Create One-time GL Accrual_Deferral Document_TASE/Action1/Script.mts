

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.03.09 Create One-time GL Accrual_Deferral Document
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

gstrTestCaseName = "Test_09.07.01.03.09 Create One-time GL Accrual_Deferral Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''''''''--------TransactionCode-FBS1----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_FBS1_0100_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date","BKPF-BUDAT", "", ConvertDate(DT_FBS1_0100_POSTING_DATE), False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBS1_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FBS1_0100_CURRENCYRATE,False)
Call SetTextbox("Type","BKPF-BLART", "", DT_FBS1_0100_TYPE, False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FBS1_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Reversal Reason","BKPF-STGRD","",DT_FBS1_0100_REVERSAL_REASON,False)
Call SetTextbox("Reversal date","BKPF-STODT", "", ConvertDate(DT_FBS1_0100_REVERSAL_DATE), False)
Call SetTextbox("PstKy","RF05A-NEWBS", "", DT_FBS1_0100_PSTKY, False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FBS1_0300_TAX_CODE,False)
Call SetTextbox("Value date","BSEG-VALUT", "", ConvertDate(DT_FBS1_0300_VALUE_DATE), False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FBS1_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FBS1_1007_COST_CENTER,False)
Call SetTextbox("PstKy","RF05A-NEWBS", "", DT_FBS1_0300_PSTKY, False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FBS1_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBS1_0300_AMOUNT_OCC1,False)
Call SetTextbox("Value date","BSEG-VALUT", "", ConvertDate(DT_FBS1_0300_VALUE_DATE_OCC1), False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FBS1_0300_TEXT,False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Wait 2
Call PressEnter()
Wait 2

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_FBS1_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FBS1_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_FBS1_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FBS1_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''''--------TransactionCode-FAGLL03----------''''
'
Call SetTcode(DT_FBS1_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL", "All Items", False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FBS1_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FBS1_1000_COMPANY_CODE,False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#4;#1")

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FBS1_0100_DOCUMENT_NUMBER,False)
Call PressEnter()
Call TakeScreenShot  
Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FBS1_0100_POSTING_DATE),False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call ClickBUtton("Back   \(F3\)",False)
Call TakeScreenShot  
Call ClickBUtton("Back   \(F3\)",False)
Call TakeScreenShot  

'''''''''--------TransactionCode-FS10N----------''''
Call SetTcode(DT_FBS1_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SO_SAKNR-LOW","",DT_FBS1_1000_GL_ACCOUNT_OCC1,False)
Call SetTextbox("Company code","SO_BUKRS-LOW","",DT_FBS1_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal year","GP_GJAHR","",Year(DT_FBS1_1000_FISCAL_YEAR),False)
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  

Call DoubleClickGuiGridCell("", 0, Cint(Month(DT_FBS1_0030_GRIDCELL_6_BALANCE)+1), "Balance", False)

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FBS1_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call FindRowNumber("", "Document Number", DT_FBS1_1105_DOCUMENT_NUMBER, "DT_ROW_OUTPUT")

Call VerifyGridCellContent("", DT_ROW_OUTPUT, "BELNR", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
Call VerifyGridCellContent("", DT_ROW_OUTPUT, "BLART", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("", DT_ROW_OUTPUT, "DMSHB", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("", DT_ROW_OUTPUT, "BLDAT", 0, ConvertDate(DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1))
Call VerifyGridCellContent("", DT_ROW_OUTPUT, "HKONT", 0, DT_FBS1_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

Call LogOff'
Call FinalStatus()

