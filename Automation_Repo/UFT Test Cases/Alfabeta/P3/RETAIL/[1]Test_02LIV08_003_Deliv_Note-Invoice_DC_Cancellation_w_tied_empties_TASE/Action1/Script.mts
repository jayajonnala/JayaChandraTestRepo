

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV08_003_Deliv_Note-Invoice_DC_Cancellation_w_tied_empties
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

gstrTestCaseName = "Test_02LIV08_003_Deliv_Note-Invoice_DC_Cancellation_w_tied_empties"
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
'
'''''''--------------------------------MR8M-----------------------------
'''
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call  SetTextbox("Invoice Document No\.","RBKPV-BELNR","",DT_MR8M_300_INVOICE_DOCUMENT_NO,False)
Call  SetTextbox("Fiscal Year","RBKPV-GJAHR","",DT_MR8M_300_FISCAL_YEAR,False)
Call  SetTextbox("Reversal Reason","UF05A-STGRD","",DT_MR8M_300_REVERSAL_REASON,False)
Call PressEnter()  
Call TakeScreenShot()
Call  ClickButton("Reverse   \(Ctrl\+S\)",False)

wait 2
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_NUM_OUTPUT")
Call VerifyStatusBar("Document reversed with no. "& DT_NUM_OUTPUT &": Please manually clear FI documents")
'


'''''--------------------------------WE09-----------------------------
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_300_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_MR8M_100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_NUM_OUTPUT,False)
 
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call GetLabelContentByRefLabel(DT_MR8M_100_VARIANT,994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call TakeScreenShot()
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call TakeScreenShot()
'Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;Segment 000002;Segment 000004","ZZMDIV_E1MBGM","",False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;#2;ZZMDIV_E1MBGM","ZZMDIV_E1MBGM","",False)
Call TakeScreenShot()

'
''''--------------------------------MIR4-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call ClickButton("Display Document   \(F2\)",False) 
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
'Call GetGridContentByTitle("Documents in Accounting",0,"Doc. Number",1,"DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call GetGridContentByTitle("Documents in Accounting",0,"Document Number",1,"DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call  ClickButtonIfExist("Cancel   \(F12\)",True)



''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MR8M_6000_OKCD)  
Call PressEnter() 
call CheckTCodeScreen(DT_MR8M_6000_OKCD)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT,False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE_OCC2,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call VerifyStatusBar("Document PU"& Right(DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT,8) &" was posted in company code GR02")


''''--------------------------------J1GP01-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTcode(DT_MR8M_6000_OKCD_1)  
Call PressEnter() 
call CheckTCodeScreen(DT_MR8M_6000_OKCD_1)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MR8M_1000_DOCUMENT_OCC2_1,False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE_OCC2_1,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)

Call VerifyStatusBar(DT_MR8M_120_CHECK_TEXT_OF_STATUSBAR_1)


''--------------------------------MIGO-----------------------------


Call SetTcode(DT_MR8M_120_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)
'
call SetComboByKey("GODYNPRO-ACTION","A03")
call SetComboByKey("GODYNPRO-REFDOC","R02")
wait 2
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MR8M_2010_GODYNPROMAT_DOC,False)
Call TakeScreenShot()
Call PressEnter()

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MR8M_0304_ITEM_OK_OCC3,False)
Call PressEnter()
Call TakeScreenShot()

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MR8M_0304_ITEM_OK_OCC3,False)
Call ClickButton("OK_NEXT_ITEM",False)
CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MR8M_0304_ITEM_OK_OCC3,False)
Call ClickButton("OK_NEXT_ITEM",False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_GR_NUM_OUTPUT")
Call VerifyStatusBar("Article document " & DT_GR_NUM_OUTPUT &" posted")


Call LogOff()
Call FinalStatus ()


