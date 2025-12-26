

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV06_002_LIV_Loc_Vend_DC_and_DSD_Subseq_Credit_for_Credit_Note
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

gstrTestCaseName = "Test_Subseq_Credit_for_Credit_Note"
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


'''--------------------------------MIRO-----------------------------

'
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,True)  
Call TakeScreenShot()
call ClickButton("Continue   \(Enter\)",True)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
call ClickButtonIfExist("Continue   \(Enter\)",True)


Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_DOCUMENT_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG",DT_MIRO_6000_TRANSACTION)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-LFSNR","",DT_MIRO_6212_RM08MLFSNR,False)

Call TakeScreenShot()
Call ClickButton("More Allocation Criteria",False)
Call TakeScreenShot()
Call ClickButton("Adopt   \(F8\)",True)
Call SelectCheckbox("INVFO-XMWST",0,"ON",False)
Call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",DT_MIRO_6310_TABLECELL_AMOUNT_0,False) 
Call SetTableData("SAPLMR1MTC_MR1M","Amount","2","","",DT_MIRO_6310_TABLECELL_AMOUNT_1,False) 
Call SetTableData("SAPLMR1MTC_MR1M","Amount","3","","",DT_MIRO_6310_TABLECELL_AMOUNT_2,False) 
Call SetTableData("SAPLMR1MTC_MR1M","Amount","4","","",DT_MIRO_6310_TABLECELL_AMOUNT_3,False) 
Call TakeScreenShot()
Call PressEnter()
Call GetTextboxValue("RM08M-DIFFERENZ","","DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)

Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT),False)
Call TakeScreenShot()
Call SelectTab("HEADER","Details",False)
Call SetComboByKey("INVFO-BLART",DT_MIRO_0150_DOC_TYPE)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBar("Document")
Call GetStatusBar("item1","DT_SUBSECUENT_CREDIT_INVOICE_NUM")

Call LogOff()
Call FinalStatus ()


