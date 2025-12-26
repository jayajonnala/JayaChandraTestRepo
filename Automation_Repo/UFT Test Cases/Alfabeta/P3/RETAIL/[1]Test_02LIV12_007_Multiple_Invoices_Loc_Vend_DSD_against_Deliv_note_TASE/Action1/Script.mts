

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV12_007_Multiple_Invoices_Loc_Vend_DSD_against_Deliv_note
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

gstrTestCaseName = "Test_02LIV12_007_DSD_against_Deliv_note"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
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
call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG","1")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE_OCC2,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)

Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call PressEnter()

Call ClickButton("Deselect All",False)
Call SelectRowGuiTableByRow("SAPLMR1MTC_MR1M",1,False)
Call SelectRowGuiTableByRow("SAPLMR1MTC_MR1M",2,False)
Call ClickButton("Delete Unselected Items",False)

Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT1",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT1),False)
Call PressEnter()
'
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_MIRO_6000_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_MIRO_6000_GET_TEXT_OF_STATUSBAR_OUTPUT & " created")

''----------------------------------------MIRO-----------------------------------------
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE_OCC2),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG","1")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE_OCC3,False)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR_OCC2,False)

Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX_OCC2,False)
Call PressEnter()

'Call ClickButton("Deselect All",False)
'Call SelectRowGuiTableByRow("SAPLMR1MTC_MR1M",3,False)
Call ClickButton("Delete Unselected Items",False)

Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT2",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT2),False)
Call PressEnter()
'
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_MIRO_6000_GET_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_MIRO_6000_GET_TEXT_OF_STATUSBAR_OCC2_OUTPUT & " created")
'
'
'''''--------------------------------WE09-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIRO_6000_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_MIRO_100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MIRO_1000_FOR_VALUE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_MIRO_100_VARIANT,Len(DT_MIRO_100_VARIANT)-1),994,0,"DT_IDOC_OUTPUT1",False)
Call ClickLabel(DT_IDOC_OUTPUT1,0,False)


Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT1&";Data records;Segment 000002","ZZMDIV_E1RSEG","",False)

Call TakeScreenShot()

'''''--------------------------------WE09-----------------------------
Call SetTcode(DT_MIRO_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_MIRO_100_OKCD)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_MIRO_100_VARIANT_OCC2,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 

Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MIRO_1000_FOR_VALUE__OCC2,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(Right(DT_MIRO_100_VARIANT_OCC2,Len(DT_MIRO_100_VARIANT_OCC2)-1),994,0,"DT_IDOC_OUTPUT2",False)
Call ClickLabel(DT_IDOC_OUTPUT2,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT2&";Data records;Segment 000002","ZZMDIV_E1RSEG","",False)
Call TakeScreenShot()

''--------------------------------MIR4-----------------------------

Call SetTcode(DT_MIRO_100_OKCD_OCC2)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(Date),False)
Call ClickButton("Display Document   \(F2\)",False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
'Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call TakeScreenShot()
Call  ClickButton("Cancel   \(F12\)",True)

'--------------------------------MIR4-----------------------------

Call SetTcode(DT_MIRO_6000_OKCD_OCC2)  
Call PressEnter()   
call CheckTCodeScreen(DT_MIRO_6000_OKCD_OCC2)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO_OCC2,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(Date),False)
Call ClickButton("Display Document   \(F2\)",False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 

Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OCC2_OUTPUT")
Call TakeScreenShot()
Call  ClickButton("Cancel   \(F12\)",True)

''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MIRO_6000_OKCD_OCC3)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT,False)
Call SetTextbox("to","S_BELNR-HIGH","",DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OCC2_OUTPUT,False)

Call SelectCheckbox("S_TEST",0,DT_MIRO_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar("posted")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyifGuiLabelExists( "Document PU"&RIGHT(DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT,8)&" was posted in company code GR02")
Call VerifyifGuiLabelExists( "Document PU"&RIGHT(DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OCC2_OUTPUT,8)&" was posted in company code GR02")

Call LogOff()
Call FinalStatus ()


