

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV12_005_LIV_Loc_Vend_Invoice_w_Prior_Vendor
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

gstrTestCaseName = "Test_02LIV12_005_LIV_Loc_Vend_Invoice_w_Prior_Vendor"
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
'''--------------------------------MIRO-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
call ClickButtonIfExist("Continue   \(Enter\)",True)

 Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG","1")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_ME21N_6020_RM08MREFERENZBELEGTYP)
Wait 5
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call PressEnter()


Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call GetStatusBar("item1","DT_DOC_NUM_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_DOC_NUM_OUTPUT & " created")

''''--------------------------------MIR4-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_MIRO_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call ClickButton("Display Document   \(F2\)",False) 
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 
Call GetGridContentByTitle("Documents in Accounting",0,"Doc. Number",1,"DT_MIRO_200_GRIDCELL_0_DOC_NUMBER_OUTPUT")
'Call GetGridContentByTitle("Documents in Accounting",0,"Document Number",1,"DT_MIRO_200_GRIDCELL_0_DOC_NUMBER_OUTPUT")
Call  ClickButtonIfExist("Cancel   \(F12\)",True)



''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_MIRO_6000_OKCD)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_200_GRIDCELL_0_DOC_NUMBER_OUTPUT,False)
Call SelectCheckbox("S_TEST",0,DT_MIRO_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_PU_NUM_OUTPUT")
Call VerifyifGuiLabelExists( "Document "&DT_PU_NUM_OUTPUT&" was posted in company code "& DT_MIRO_1000_COMPANY_CODE)


Call LogOff()
Call FinalStatus ()


