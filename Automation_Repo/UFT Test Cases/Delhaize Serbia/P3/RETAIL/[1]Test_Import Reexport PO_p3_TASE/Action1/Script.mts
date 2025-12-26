'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Import Reexport PO_p3
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

gstrTestCaseName = "Test_Import Reexport PO_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear AP Accounts - Manual_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-ME9F----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIR4_6150_INVOICE_DOCUMENT_NO,False)  
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIR4_6150_FISCAL_YEAR,False) 
Call TakeScreenShot
Call ClickBUtton("Display Document   \(F2\)",False)
Call TakeScreenShot
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BSCHL", "", DT_MIR4_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", "", DT_MIR4_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "AZBET", "", DT_MIR4_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyTextBoxContent("Posting Date","BKPF-BUDAT", "", ConvertDate(DT_MIR4_0750_CHECK_TEXT_OF_POSTING_DATE), False)

''''--------TransactionCode-MIGO----------''''
Call SetTcode(DT_MIR4_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetComboByKey("GODYNPRO-ACTION", DT_MIR4_0010_GODYNPROACTION)

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER","",DT_MIR4_2000_GODYNPROPO_NUMBER, False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIR4_0110_DELIVERY_NOTE,False)  
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("GODYNPRO-DETAIL_TAKE", DT_MIR4_0200_CHECK_SELECTED_OF_TABLECELL_OK_0)
Call CLickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
Call GetTextStatusBar("DT_MIR4_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MIR4_0001_CHECK_TEXT_OF_STATUSBAR_OUTPUT",DT_MIR4_0001_CHECK_TEXT_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIR4_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call SetComboByKey("GODYNPRO-ACTION", DT_MIR4_0010_GODYNPROACTION_OCC1)
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("TS_GOHEAD","Doc. info",False)
Call TakeScreenShot
Call ClickButton("FI Documents",False)
Call TakeScreenShot

'*--------------- Uncomment below lines if Documents in Accounting grid pop up is coming  -------------------*'
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)

Call SelectColumnGuiGrid("", "", "Amount", False)
Call ClickButtonToolBar("&MB_SUM",0)

Call VerifyGridCellContent("", 4, "Amount", "", DT_POSITIVE_OCC1)
Call LogOff
Call FinalStatus()
