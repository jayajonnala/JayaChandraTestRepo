'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_FB65-FBL1N-Post credit note
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_FB65-FBL1N-Post credit note"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''--------TransactionCode-FB65----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call PressEnter()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_FB60_0010_POSTING_DATE),False)
Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB60_0010_TAX_AMOUNT,False)
Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
'this function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Amount","INVFO-WAERS","",DT_FB60_0010_AMOUNT_OCC1,False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE,False)
Call SetTextbox("Document date","INVFO-BLDAT","",ConvertDate(DT_FB60_0010_INVOICE_DATE),False)
Call TakeScreenShot
Call SetTableData("SAPLFSKBTABLE", "G/L acct", 1, "", "",DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", 1, "", "",DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", 1, "", "",DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call SetTableData("SAPLFSKBTABLE", "Tax code", 1, "", "",DT_FB60_0100_TABLECELL_TAX_CODE_0, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call VerifyTextBoxContent("Bal.", "RF05A-AZSAL", "", DT_FB60_1100_CHECK_TOOLTIP_OF_BAL, False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_OP_DOC_NO1")
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(DT_FB60_1000_CHECK_TEXT_OF_COMPANY_CODE)
Call ClickButtonIfExist("Cancel   \(F12\)",False)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
'
'''''''--------TransactionCode-/fbl1n----------''''

Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SelectRadioButton("X_OPSEL", "Open items", False)
Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FB60_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB60_1000_VENDOR_ACCOUNT,False)
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN012-LOW","",DT_FB60_1106_DOCUMENT_NUMBER,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot
Call SelectMenuBar("Settings;Switch list")
Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
' SelectRowGuiGrid(gridTitle, gridIndex, columnName, refValue, blnIsItPopup)
'Call SelectRowGuiGrid("Displayed Columns","","Column Name","Clearing Document",True)
'Call SelectRowGuiGrid("Displayed Columns","","Column Name","Text",True)
'Call ClickButtonIfExist("APP_FL_SING",True)
''Call SelectRowGuiTable("SAPLSKBHTC_WRITE_LIST","Column content","Clearing Document",True)
''Call SelectRowGuiTable("SAPLSKBHTC_WRITE_LIST","Column content","Text",True)
''Call ClickButtonIfExist("Hide Sel\. Fields \(Ctrl\+F2\)",True)
''Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST","Content","Posting Date",True)
''Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST","Content","Net due date",True)
''Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST","Content","Terms of payment",True)
''Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST","Content","Invoice reference",True)
'Call SelectRowGuiGrid("Column Set","","Column Name","Net due date",True)
'Call SelectRowGuiGrid("Column Set","","Column Name","Terms of Payment",True)
'Call SelectRowGuiGrid("Column Set","","Column Name","Posting Date",True)
Call SelectRowGuiGrid("Column Set","","Column Name","Tax code",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call SelectRowGuiGrid("Column Set","","Column Name","Invoice reference",True)
'Call ClickButtonIfExist("APP_FL_SING",True)
'Call SelectRowGuiTable("SAPLSKBHTC_FIELD_LIST","Content","Tax code",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
''Call ClickButtonIfExist("Copy   \(Enter\)",True)
'' VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContentbyName("shell",1,"Document Number","",DT_FB60_1106_DOCUMENT_NUMBER)
Call VerifyGridCellContentbyName("shell",1,"Document Date","",ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
'Call VerifyGridCellContentbyName("shell",1,"Reference","",DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR)
Call VerifyGridCellContentbyName("shell",1,"Amount in local currency","",DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call SetHorizontalScrollBar(130,False)
Call VerifyGridCellContentbyName("shell",1,"Net due date","",ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FAEDT))
Call VerifyGridCellContentbyName("shell",1,"Tax code","",DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContentbyName("shell",1,"Terms of Payment","",DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_1106_DOCUMENT_NUMBER,0)
'''Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT),0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XBLNR,0)
'''Call VerifyifGuiLabelExists_ByIndex(ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FAEDT),0)
''' SetHorizontalScrollBar(ScrollBarPosition, blnIsItPopup)
'''Call SetHorizontalScrollBar(130,False)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM,0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ,0)
'''Call VerifyifGuiLabelExists_ByIndex(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,0)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
