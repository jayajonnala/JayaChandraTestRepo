

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV10_001_Subsequent_Debit_Cancellation_against_Invoice
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

gstrTestCaseName = "Test_02LIV10_001_Subsequent_Debit_Cancellation_against_Invoice"
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
''''''''--------------------------------MR8M-----------------------------
''''
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call  SetTextbox("Invoice Document No\.","RBKPV-BELNR","",DT_MR8M_300_INVOICE_DOCUMENT_NO,False)
Call  SetTextbox("Fiscal Year","RBKPV-GJAHR","",Year(Date),False)
Call  SetTextbox("Reversal Reason","UF05A-STGRD","",DT_MR8M_300_REVERSAL_REASON,False)
Call  SetTextbox("Posting Date","G_BUDAT","",ConvertDate(DT_MR8M_300_POSTING_DATE),False)

Call ClickButton("Reverse   \(Ctrl\+S\)",False)

wait 2
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_DOCUMENT_REVERSED_NUMBER_OUTPUT")
Call VerifyStatusBar("Document reversed with no. "& DT_DOCUMENT_REVERSED_NUMBER_OUTPUT &": Please manually clear FI documents")

''
'''
''''''''--------------------------------WE09-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_300_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9"&".*","","Variant name",DT_MR8M_600_GRIDCELL_5_VARIANT_NAME,True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MR8M_1000_FOR_VALUE,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBar("IDocs were found")

Call VerifyifGuiLabelExists(DT_MR8M_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)
Call GetLabelContentByRefLabel(Right(DT_MR8M_600_GRIDCELL_5_VARIANT_NAME,Len(DT_MR8M_600_GRIDCELL_5_VARIANT_NAME)-1),994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call TakeScreenShot()
'
'
'''''--------------------------------MIR4-----------------------------

Call SetTcode(DT_MR8M_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)


Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MR8M_6150_INVOICE_DOCUMENT_NO,False) 
Call ClickButton("Display Document   \(F2\)",False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
'Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call TakeScreenShot()
Call  ClickButtonIfExist("Cancel   \(F12\)",True)
'
'
'''''--------------------------------J1GP01-----------------------------
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MR8M_6000_OKCD)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTextbox("Company Code","S_BUKRS","",DT_MR8M_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","S_GJAHR","",Year(Date),False)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MR8M_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT,False)
Call SelectCheckbox("S_TEST",0,DT_MR8M_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()


Call LogOff()
Call FinalStatus ()



