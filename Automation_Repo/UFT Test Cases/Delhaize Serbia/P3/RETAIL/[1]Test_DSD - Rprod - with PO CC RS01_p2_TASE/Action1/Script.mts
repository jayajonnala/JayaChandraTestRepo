'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Reload DataSheet to updates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_DSD - Rprod - with PO CC RS01_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 13th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_DSD - Rprod - with PO CC RS01_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_DSD - Rprod - with PO CC RS01_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

''Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_MIGO_0010_REFERENCE",(Cint(DT_MIGO_0010_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_2010_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD",DT_MIGO_0100_DOC_INFO,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("FI Documents",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'___________________________________________________enable if required________________________________________________________
'Call SelectRowGuiGridbyRowNo("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER,true)
''Capture the screenshot
'Call TakeScreenShot()
'Call ClickButton("Display Document   \(F2\)",True)
'Wait(1)
''Capture the screenshot
'Call TakeScreenShot()
''__________________________________________________________end block_________________________________________________________
'
Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'
'-------------------------------------------------MIR7-----------------------------------------
Call SetTcode(DT_MIGO_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0100_OKCD)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIGO_1000_COMPANY_CODE)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetComboByKey("RM08M-VORGANG","1")

'Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace(DT_MIGO_0010_INVOICE_DATE,"/","."),False)

Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIGO_0010_INVOICE_DATE),False)
'Call PressEnter() 
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
wait(1)
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_MIGO_0010_REFERENCE,False)
wait(1)
Call PressEnter()

'Enter the delivery Note No
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIGO_6020_RM08MREFERENZBELEGTYP)
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIGO_6211_RM08MEBELN,False)
'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_MIGO_0010_CALCULATE_TAX,False)
Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Click on Post Buton
Call ClickButtonIfExist("Save Parked Document   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_MIGO_INVOICE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_STATUS_BAR_INVOICE)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'
'-----------------------------------------------zmdiv_invoice_list-----------------------------------------------------
Call SetTcode(DT_MIGO_0100_OKCD_OCC1) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIGO_0100_OKCD_OCC1)
Call TakeScreenShot()
'Display the Invoice Details
'Call SetTextbox("Company Code","S_BUKRS-LOW","",DT_MIGO_1000_COMPANY_CODE_OCC1,False)
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot()
Call SelectRowGuiGrid("Variant Catalog for Program ZMDIV_POST_PARKED_INV_DS.*", "", "Variant Name", "PDJOKO3", True)
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",True)
Call TakeScreenShot()
Call SetTextbox("Proposed Act\. Gds Mvmnt Date","P_WADAT","",Replace(DT_MIGO_1000_PROPOSED_ACT_GDS_MVMNT_DATE,"/","."),False)
Call SetTextbox("Document Date","S_BLDAT-LOW","",Replace(DT_MIGO_1000_PROPOSED_ACT_GDS_MVMNT_DATE,"/","."),False)
Call SetTextbox("to","S_BLDAT-HIGH","",Replace(DT_MIGO_1000_PROPOSED_ACT_GDS_MVMNT_DATE,"/","."),False)
Call TakeScreenShot()
Call PressEnter()
wait(1)
Call TakeScreenShot()
'Click on Follow On Document
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(1)
Call TakeScreenShot()

'Click on Follow On Document
Call ClickButtonIfExist("Invoice Status",False)
wait(1)
Call TakeScreenShot()

'Call SelectCheckbox("P_DLV",0,DT_MIGO_1000_INBOUND_DELIVERIES,False)

'Click on execute
Call ClickButtonIfExist("Execute   \(F8\)",False)
wait(1)
Call TakeScreenShot()
'filter wit doc nummber
Call SelectColumnGuiGrid("","","Document Number",False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+F5\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_MIGO_1000_INVOICE_DOCUMENT_NO,True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(Enter\)",True)
'select the row
Call SelectRowGuiGridbyRowNo("","",DT_SET_SELECTED_ROW,False)
Call ClickCellGuiGrid("", "", "Document Number", 1, "", "", False)
Call TakeScreenShot()
'Click on post
Call ClickButtonIfExist("Post   \(Shift\+F11\)",False)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(1)
Call TakeScreenShot()

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


