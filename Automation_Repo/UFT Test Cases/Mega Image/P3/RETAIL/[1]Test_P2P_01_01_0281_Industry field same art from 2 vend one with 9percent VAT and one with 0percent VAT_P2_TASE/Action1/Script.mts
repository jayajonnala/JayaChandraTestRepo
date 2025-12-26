
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0281_Industry field same art from 2 vend one with 9percent VAT and one with 0percent VAT_P2
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_P2P_01_01_0281_one with 0percent VAT_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'//SetTransactionCode- MIRO//

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code","RO02")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetComboByKey("RM08M-VORGANG", DT_MIRO_0010_TRANSACTION_KEY)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call SetTextbox("Posting date","INVFO-BUDAT","",ConvertDate(DT_MIRO_0010_POSTING_DATE),False)
Call PressEnter()

Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_MIRO_6211_RM08MEBELN, False)
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6211_RM08MXWARE_BNK)
If VerifyTextBoxEnabled("Posting date","INVFO-BUDAT", 0, False)<>True Then
	Call PressEnter()
End If 
Call SetComboByKey("RM08M-XWARE_BNK", DT_MIRO_6211_RM08MXWARE_BNK)
Call PressEnter()
Call SelectCheckbox("INVFO-XMWST", 0, DT_MIRO_0010_CALCULATE_TAX, False)
Call TakeScreenShot

Call GetTextboxValue("INVFO-WMWST", "", "DT_MIRO_0010_CHECK_TEXT_OF_TAX_AMOUNT_OUTPUT", false)
Call GetTextboxValue("RM08M-DIFFERENZ", "", "DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", false)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT),False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",false)

Call GetStatusBar("item1","DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Document no. "&DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT& " created")

Call SetTcode(DT_SAPTRANSACTIONCODE_DISPLAY_DOC) 
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Display Document   \(F2\)",false)

Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)

Call TakeScreenShot

Call GetTextboxValue("BKPF-BELNR", "", "DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", false)

Call VerifyGridCellContentbyName("shell", 1, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContentbyName("shell", 2, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContentbyName("shell", 3, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContentbyName("shell", 1, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContentbyName("shell", 2, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContentbyName("shell", 3, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

''Call VerifyGridCellContentbyName("shell", 3, "Amount", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContentbyName("shell", 3, "Amount", "", DT_MIRO_0010_CHECK_TEXT_OF_TAX_AMOUNT_OUTPUT)
''''--------2nd Invoice
''//SetTransactionCode- MIRO//

Call SetTcode(DT_MIRO_0750_OKCD)     
Call PressEnter()     

Call SetComboByKey("RM08M-VORGANG", DT_MIRO_6211_RM08MXWARE_BNK_OCC1)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE_OCC1),False)
Call PressEnter()
''Call SetTextbox("Posting date","INVFO-BUDAT","",ConvertDate(DT_MIRO_0010_POSTING_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE_OCC1,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_MIRO_6211_RM08MEBELN_OCC1, False)
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6211_RM08MXWARE_BNK)
Call PressEnter()
Call SelectCheckbox("INVFO-XMWST", 0, DT_MIRO_0010_CALCULATE_TAX_OCC1, False)
Call TakeScreenShot

Call GetTextboxValue("INVFO-WMWST", "", "DT_MIRO_0010_CHECK_TEXT_OF_TAX_AMOUNT_OCC1_OUTPUT", false)
Call GetTextboxValue("RM08M-DIFFERENZ", "", "DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OCC1_OUTPUT", false)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OCC1_OUTPUT),False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",false)

Call GetStatusBar("item1","DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT& " created")

Call SetTcode(DT_SAPTRANSACTIONCODE_DISPLAY_DOC) 
Call PressEnter()

Call TakeScreenShot
Call ClickButton("Display Document   \(F2\)",false)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)
Call PressEnter()
Call TakeScreenShot

Call GetTextboxValue("BKPF-BELNR", "", "DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT", false)

Call VerifyGridCellContentbyName("shell", 1, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContentbyName("shell", 2, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
''Call VerifyGridCellContentbyName("shell", 3, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContentbyName("shell", 1, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContentbyName("shell", 2, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
'
''''------TCode-------------

Call SetTcode(DT_MIRO_0750_OKCD_OCC1)     
Call PressEnter()

Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_MIRO_1000_COMPANY_CODE,False)
'Call SetTextbox("Document Number","BR_BELNR-LOW","","",False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","","2020",False)

Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_MIRO_1000_POSTING_DATE),False)

Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH",false)

Call SetTableData("SAPLALDBSINGLE","Single value",1,"","","5200451682",True)
Call SetTableData("SAPLALDBSINGLE","Single value",2,"","","5200451683",True)

Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",false)




Call LogOff()

Call FinalStatus ()


