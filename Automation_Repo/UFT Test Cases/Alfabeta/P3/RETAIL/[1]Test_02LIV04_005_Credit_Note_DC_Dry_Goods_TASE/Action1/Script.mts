

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV04_005_Credit_Note_DC_Dry_Goods
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

gstrTestCaseName = "Test_02LIV04_005_Credit_Note_DC_Dry_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''


''--------------------------------MIRO-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
call ClickButtonIfExist("Continue   \(Enter\)",True)

 Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_DOCUMENT_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG","2")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP_OCC1)
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call PressEnter()
Call TakeScreenShot()

Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)
Call SelectTab("HEADER","Details",False)
Call SetComboByKey("INVFO-BLART",DT_MIRO_0150_DOC_TYPE)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call GetStatusBar("item1","DT_NUM_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_NUM_OUTPUT & " created")


'''''--------------------------------WE09-----------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MIRO_6000_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Basic Type","IDOCTP-LOW","",DT_MIRO_1000_BASIC_TYPE,False)
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_MIRO_1000_LOGICAL_MESSAGE,False)
Call SetTextbox("Search in Segment \.\.\.","SEGMENT1","",DT_MIRO_1000_SEARCH_IN_SEGMENT,False) 
Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_MIRO_1000_SEARCH_IN_FIELD,False) 

Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MIRO_1000_FOR_VALUE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_MIRO_1000_LOGICAL_MESSAGE,Len(DT_MIRO_1000_LOGICAL_MESSAGE)),994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","BELNR","Fld Cont.",DT_MIRO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)
DT_MIRO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_Calculate=Year(Date) & ConvertDoubledigit(CSTR(Month(Date))) & ConvertDoubledigit(CSTR(Day(Date)))
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","BLDAT","Fld Cont.",DT_MIRO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_Calculate)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call TakeScreenShot()



'''''--------------------------------MIR4-----------------------------
Call SetTcode(DT_MIRO_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(Date()),False)
Call ClickButton("Display Document   \(F2\)",False)

Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 

Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
'Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
Call TakeScreenShot()
Call  ClickButton("Cancel   \(F12\)",True)

'''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MIRO_6000_OKCD_OCC2)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)


Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT,False)

Call SelectCheckbox("S_TEST",0,DT_MIRO_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar("posted")
Call GetStatusBar("item1","DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT")
Call VerifyifGuiLabelExists( "Document "&DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT&" was posted in company code "& DT_MIRO_1000_COMPANY_CODE)


Call LogOff()
Call FinalStatus ()



