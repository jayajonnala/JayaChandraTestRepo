'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Import Reexport PO_p2
'.................Author : TCS        
'................ Creation Date    :
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

gstrTestCaseName = "Test_Import Reexport PO_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-ME9F----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REF",Cint(DT_INCREMENT_REF)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Document Number","S_EBELN-LOW","",DT_ME9F_1000_DOCUMENT_NUMBER,False)  
Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME9F_1000_PURCHASING_ORGANIZATION,False)  
Call SetTextbox("Application","P_KAPPL","",DT_ME9F_1000_APPLICATION,False)  
Call SetTextbox("Processing Status","P_VSTAT","",DT_ME9F_1000_PROCESSING_STATUS,False)  
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SelectCheckboxNoLabel(0, DT_ME9F_0120_NO_NAME, False)
Call ClickBUtton("Display Message   \(Shift\+F8\)",False)
Call TakeScreenShot

''''--------TransactionCode-MIRO----------''''
Call SetTcode(DT_ME9F_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_ME9F_1000_COMPANY_CODE)
Call ClickBUttonifExist("Continue   \(Enter\)",True)

Call SetComboByKey("RM08M-VORGANG", DT_INVOICE_PICKER)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_ME9F_0010_INVOICE_DATE),False) 
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Reference","INVFO-XBLNR","",DT_ME9F_0010_REFERENCE,False)  
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_ME9F_0010_REFERENCE,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WAERS","","RSD",False) 
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_ME9F_6211_RM08MXWARE_BNK)
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_ME9F_6211_RM08MEBELN, False)
Call PressEnter()     
Call TakeScreenShot
Call SetTableData("SAPLMR1MTC_MR1M","Amount",1,"","",DT_ME9F_6310_TABLECELL_AMOUNT_0,False)
Call SetTableData("SAPLMR1MTC_MR1M","Quantity",1,"","",DT_ME9F_6310_TABLECELL_QUANTITY_0,False)

Call SelectCheckbox("INVFO-XMWST",0,DT_ME9F_0010_CALCULATE_TAX,False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",1,"ZZ (0% Procurement - no VAT relevant)",False)

Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WRBTR","",DT_ME9F_0010_AMOUNT,False) 
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetTextStatusBar("DT_ME9F_6000_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME9F_6000_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT",DT_ME9F_6000_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call verifyStatusBar(DT_ME9F_6000_CHECK_TEXT_OF_STATUSBAR_OCC2)
Call LogOff'




