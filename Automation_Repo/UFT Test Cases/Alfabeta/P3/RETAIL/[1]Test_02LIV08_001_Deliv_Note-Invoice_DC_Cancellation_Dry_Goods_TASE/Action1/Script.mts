

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV08_001_Deliv_Note-Invoice_DC_Cancellation_Dry_Goods
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

gstrTestCaseName = "Test_02LIV08_001_Deliv_Note-Invoice_DC_Cancellation_Dry_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''
'''''--------------------------------MR8M-----------------------------
''
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call  SetTextbox("Invoice Document No\.","RBKPV-BELNR","",DT_MR8M_300_INVOICE_DOCUMENT_NO,False)
Call  SetTextbox("Fiscal Year","RBKPV-GJAHR","",DT_MR8M_300_FISCAL_YEAR,False)
Call  SetTextbox("Reversal Reason","UF05A-STGRD","",DT_MR8M_300_REVERSAL_REASON,False)
Call SelectMenuBar("Invoice Document;Reverse")
wait 2
Call GetStatusBar("item1","DT_NUM_OUTPUT")
Call VerifyStatusBar("Document reversed with no. "& DT_NUM_OUTPUT &": Please manually clear FI documents")
''
'
'
'''''''--------------------------------WE09-----------------------------
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_300_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'
 Call SetTextbox("Direction \(1=Outb\., 2=Inb\.\)","DIRECT-LOW","",DT_MR8M_1000_DIRECTION__1OUTB__2INB,False)
 Call SetTextbox("Basic Type","IDOCTP-LOW","",DT_MR8M_1000_BASIC_TYPE,False)
 Call SetTextbox("Logical Message","MESTYP-LOW","",DT_MR8M_1000_LOGICAL_MESSAGE,False)
 Call SetTextbox("Search in Segment \.\.\.","SEGMENT1","",DT_MR8M_1000_SEARCH_IN_SEGMENT,False)
 Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_MR8M_1000_SEARCH_IN_FIELD,False)
 Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_NUM_OUTPUT,False)
 
 
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBar("IDocs were found")

Call GetLabelContentByRefLabel(DT_MR8M_1000_LOGICAL_MESSAGE,994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","BELNR","Fld Cont.",DT_NUM_OUTPUT)

'''''--------------------------------MIR4-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTcode(DT_MR8M_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MR8M_6150_INVOICE_DOCUMENT_NO,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MR8M_6150_FISCAL_YEAR,False)
Call ClickButton("Display Document   \(F2\)",False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
Call SelectRowGuiGrid("Documents in Accounting","","Object type text","Accounting document",True)
Call DoubleClickGuiGridCell("Documents in Accounting","",1,"Object type text",True)
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call  GetTextboxValue("BKPF-BELNR","","DT_PU_ORD_OUTPUT",False)
Call TakeScreenShot()
Call  ClickButton("Back   \(F3\)",False)
Call  ClickButtonIfExist("Cancel   \(F12\)",True)

'
''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MR8M_750_OKCD)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

Call SetTextbox("Document","S_BELNR-LOW","",CSTR(DT_PU_ORD_OUTPUT-1),False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar("posted")
Call VerifyifGuiLabelExists( "Document PU"&Mid(Cstr(DT_PU_ORD_OUTPUT-1),3,9)& " was posted in company code GR02")

Call SetTcode(DT_MR8M_750_OKCD2)  
Call PressEnter() 
call CheckTCodeScreen(DT_MR8M_750_OKCD2)

Call SetTextbox("Document","S_BELNR-LOW","",DT_PU_ORD_OUTPUT,False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE_2,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar("posted")
Call VerifyifGuiLabelExists( "Document PU"&Mid(DT_PU_ORD_OUTPUT,3,9)& " was posted in company code GR02")


''--------------------------------MIGO-----------------------------

'
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_120_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)
'
call SetComboByKey("GODYNPRO-ACTION",DT_MR8M_2010_TRANSACTION)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MR8M_2010_GODYNPROMAT_DOC,False)
Call PressEnter() 
call SetComboByKey("GOHEAD-WEVER",DT_MR8M_0200_SAPLMIGOTV_GOITEM)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MR8M_0304_ITEM_OK,False)

Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBar("posted")

Call LogOff()

Call FinalStatus ()



