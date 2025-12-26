

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV09_001_Credit_Note_DC_Deficit_Cancellation
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

gstrTestCaseName = "Test_02LIV09_001_Credit_Note_DC_Deficit_Cancellation"
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
'''
''
''''''''--------------------------------MR8M-----------------------------
''''
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call  SetTextbox("Invoice Document No\.","RBKPV-BELNR","",DT_MR8M_300_INVOICE_DOCUMENT_NO,False)
Call  SetTextbox("Fiscal Year","RBKPV-GJAHR","",Year(Date),False)
Call  SetTextbox("Reversal Reason","UF05A-STGRD","",DT_MR8M_300_REVERSAL_REASON,False)
Call PressEnter()  
Call TakeScreenShot()
Call  ClickButton("Reverse   \(Ctrl\+S\)",False)

wait 2
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickButton("Display Document   \(F2\)",False)
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call VerifyTextBoxNoLabelContent("RBKPV-BELNR",0,DT_MR8M_6000_CHECK_TEXT_OF_TRANSACTION,False)
Call GetWindowValue("DT_MR8M_6000_GET_TEXT_OF_TITL_OUTPUT",False)
CAll VerifyWindowValue(DT_MR8M_6000_GET_TEXT_OF_TITL_OUTPUT)

''''--------------------------------MIR4-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_6000_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_MR8M_6000_OKCD)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MR8M_6150_INVOICE_DOCUMENT_NO,False) 
Call PressEnter()   
Call VerifyTextBoxNoLabelContent("RBKPV-BELNR",0,DT_MR8M_6000_CHECK_TEXT_OF_TRANSACTION_OCC2,False)

Call SelectTab("HEADER","Details",False)
Call TakeScreenShot()

'

'''''--------------------------------WE09-----------------------------
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_6000_OKCD_OCC2)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

 Call SetTextbox("Direction \(1=Outb\., 2=Inb\.\)","DIRECT-LOW","",DT_MR8M_1000_DIRECTION__1OUTB__2INB,False)
Call SetTextbox("Basic Type","IDOCTP-LOW","",DT_MR8M_1000_BASIC_TYPE,False)
 Call SetTextbox("Logical Message","MESTYP-LOW","",DT_MR8M_1000_LOGICAL_MESSAGE,False)
 Call SetTextbox("Search in Segment \.\.\.","SEGMENT1","",DT_MR8M_1000_SEARCH_IN_SEGMENT,False)
 Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_MR8M_1000_SEARCH_IN_FIELD,False)
 Call SetTextbox("to","STATUS-HIGH","",DT_STATUS,False) 
 Call TakeScreenShot()
 Call SetTextbox("for Value \.\.\.","VALUE1_1",0,DT_MR8M_1000_FOR_VALUE,False)
  
 Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar("IDocs were found")
'Call GetLabelContentByRefLabel(DT_MR8M_1000_LOGICAL_MESSAGE,994,0,"DT_IDOC_OUTPUT",False)
Call GetLabelContentByRefLabel(DT_MR8M_1000_LOGICAL_MESSAGE,1136,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","BELNR","Fld Cont.",DT_MR8M_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)

'
'''--------------------------------MIR4-----------------------------

Call SetTcode(DT_MR8M_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MR8M_6150_INVOICE_DOCUMENT_NO_OCC2,False) 

Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(Date),False)
Call ClickButton("Display Document   \(F2\)",False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 

Call SelectRowGuiGridbyRowNo("Documents in Accounting",0,1,True)
Call DoubleClickGuiGridCell("Documents in Accounting","",1,"Object type text",True)
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
Call  GetTextboxValue("BKPF-BELNR","","DT_PU_ORD_OUTPUT",False)
Call TakeScreenShot()

''
''''--------------------------------J1GP01-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTcode(DT_MR8M_750_OKCD)  
Call PressEnter() 
call CheckTCodeScreen(DT_MR8M_750_OKCD)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MR8M_1000_DOCUMENT,False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE,False)

Call  ClickButton("Execute   \(F8\)",False)

Call SetTextbox("Document","S_BELNR-LOW","",DT_MR8M_1000_DOCUMENT_OCC2,False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE_OCC5,False)

Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")
Call VerifyifGuiLabelExists( DT_MR8M_120_CHECK_TEXT_OF_PU70000507)

Call LogOff()
Call FinalStatus ()




