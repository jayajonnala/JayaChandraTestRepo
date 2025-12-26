

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV06_001_LIV_Loc_Vend_DC_and_DSD_Subsequent_Credit_for_Invoice
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

gstrTestCaseName = "Test_02LIV06_001_LIV_for_Invoice"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''


'''''--------------------------------MIRO-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG",DT_MIRO_6000_TRANSACTION_OCC2)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-LFSNR","",DT_MIRO_6212_RM08MLFSNR,False)
Call TakeScreenShot()

Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX_OCC2,False)
Call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_0,False) 
Call SetTableData("SAPLMR1MTC_MR1M","Amount","2","","",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_1,False) 
Call SetTableData("SAPLMR1MTC_MR1M","Amount","3","","",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_2,False) 
Call SetTableData("SAPLMR1MTC_MR1M","Amount","4","","",DT_MIRO_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_3,False) 
Call TakeScreenShot()
Call PressEnter()
Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_BALANCE_OUTPUT",False)

Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_BALANCE_OUTPUT),False)
Call TakeScreenShot()

Call VerifyTableCellContent(1,"Purchase Order","SAPLMR1MTC_MR1M",DT_MIRO_6310_CHECK_TEXT_PURCHASE_ORDER_0)
Call SelectTab("HEADER","Details",False)
Call SetComboByKey("INVFO-BLART",DT_MIRO_0150_DOC_TYPE_OCC2)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBar("Document")
Call GetStatusBar("item1","DT_MIRO_6000_INV_NUMBER_OUTPUT")

Call ClickButton("Other Invoice Document   \(Ctrl\+F10\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'''''--------------------------------WE09-----------------------------
'
Call SetTcode(DT_MIRO_6000_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name",DT_MIRO_100_VARIANT,True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MIRO_6000_INV_NUMBER_OUTPUT,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_MIRO_100_VARIANT,Len(DT_MIRO_100_VARIANT)-1),994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","EBELN","Fld Cont.",DT_MIRO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_3)

'''''--------------------------------MIR4-----------------------------
Call SetTcode(DT_MIRO_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call ClickButton("Display Document   \(F2\)",False) 
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
'Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
Call TakeScreenShot()
Call  ClickButton("Cancel   \(F12\)",True)

'''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MIRO_6000_OKCD_OCC2)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT,False)
Call SelectCheckbox("S_TEST",0,DT_MIRO_1000_TEST_MODE_OCC2,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar("posted")
Call GetStatusBar("item1","DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT")
Call VerifyifGuiLabelExists( "Document "&DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT&" was posted in company code "& DT_MIRO_1000_COMPANY_CODE)

Call LogOff()
Call FinalStatus ()





