'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Import Reexport PO_p4
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

gstrTestCaseName = "Test_Import Reexport PO_p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-MIRO----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextboxPopupIfExist("BKPF-BUKRS", "Company Code", DT_MIRO_1000_COMPANY_CODE)
Call ClickBUttonifExist("Continue   \(Enter\)",True)

Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False) 
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)  
Call SetTextbox("Amount","INVFO-WAERS","",DT_MIRO_0010_AMOUNT,False)
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_MIRO_0010_REFERENCE,False)
Call SetComboByKey("RM08M-REFERENZBELEGTYP","1")
Call SetComboByKey("RM08M-XWARE_BNK", DT_MIRO_6211_RM08MXWARE_BNK)
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_MIRO_6211_RM08MEBELN, False)
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)

Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",1,"QG (20% Procur. - domestic suppliers)",False)
Call SetTableData("SAPLMR1MTC_MR1M", "Amount", 1, "", "", "6000", False)
Call SetTableData("SAPLMR1MTC_MR1M", "Quantity", 1, "", "", "15", False)

Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
 
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetTextStatusBar("DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call verifyStatusBar(DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR_OCC1)

''''--------TransactionCode-MIRO----------''''
Call SetTcode(DT_MIRO_6000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO,False)  
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIRO_6150_FISCAL_YEAR,False)

Call ClickButton("Display Document   \(F2\)",False)
Call TakeScreenShot
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False)
Call TakeScreenShot

'* Uncomment below lines if Documents in Accounting grid pop up comes up  *'

Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)

Call SelectColumnGuiGrid("", "", "Amount", False)
Call ClickButtonToolBar("&MB_SUM",0)
Call VerifyGridCellContent("",4,"Amount", "", "0,00")

Call LogOff
Call FinalStatus()


