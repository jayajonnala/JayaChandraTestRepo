

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV04_002_Credit_Note_DC_Fresh_Goods
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


gstrTestCaseName = "Test_02LIV04_002_Credit_Note_DC_Fresh_Goods"
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
Call SetComboByKey("RM08M-REFERENZBELEGTYP","2")

Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call PressEnter()

Call PressEnter()
Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)
Call SelectTab("HEADER","Details",False)
Call SetComboByKey("INVFO-BLART",DT_MIRO_0150_DOC_TYPE)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call GetStatusBar("item1","DT_DOC_NUM_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_DOC_NUM_OUTPUT & " created")


'''''--------------------------------WE09-----------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MIRO_0100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
'Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RSEIDOC9",0,DT_MIRO_0600_GRIDCELL_25_VARIANT_NAME,False)
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name",DT_MIRO_0600_GRIDCELL_25_VARIANT_NAME_OCC1,True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_MIRO_1000_CREATED_ON),False)
Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_MIRO_1000_TO),False) 

Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MIRO_1000_FOR_VALUE__OCC1,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_MIRO_0600_GRIDCELL_25_VARIANT_NAME_OCC1,Len(DT_MIRO_0600_GRIDCELL_25_VARIANT_NAME_OCC1)-1),994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call TakeScreenShot()



'''''--------------------------------MIR4-----------------------------
Call SetTcode(DT_MIRO_0100_OKCD_OCC1)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(Date()),False)
Call ClickButton("Display Document   \(F2\)",False)

Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
'Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
Call TakeScreenShot()
Call  ClickButton("Cancel   \(F12\)",True)

'''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MIRO_0100_OKCD_OCC2)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SetTextbox("Company Code","S_BUKRS","",DT_MIRO_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","S_GJAHR","",DT_MIRO_1000_FISCAL_YEAR,False)

Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT,False)
Call SetTextbox("to","S_BELNR-HIGH","",DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT,False)
Call SelectCheckbox("S_TEST",0,DT_MIRO_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar("posted")
Call GetStatusBar("item1","DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT")
Call VerifyifGuiLabelExists( "Document "&DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT&" was posted in company code "& DT_MIRO_1000_COMPANY_CODE)


Call LogOff()
Call FinalStatus ()



