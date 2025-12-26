

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV01_001_LocVend_DC_and_DSD_Invoice_Dry_Goods
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

gstrTestCaseName = "Test_02LIV01_001_LocVend_DC_and_DSD_Invoice_Dry_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


'''
'''--------------------------------MIRO-----------------------------
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,True)
 Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",TRue) 
 Call TakeScreenShot()

 Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
 Call SetTextboxNoLabel("INVFO-BUDAT","",ConvertDate(DT_MIRO_0010_POSTING_DATE),False) '------ Datelabel changes as per transaction hence attached text not used

Call SetCombo("RM08M-VORGANG","Invoice")

Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
 Call TakeScreenShot()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP_OCC6)
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
 Call TakeScreenShot()
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX_OCC2,False)
Call PressEnter()

Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call GetStatusBar("item1","DT_MIRO_6000_INV_NUMBER_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_MIRO_6000_INV_NUMBER_OUTPUT & " created")


Call LogOff()
Call FinalStatus ()






