

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_3_1_19 LIV Local Vendor DC and DSD Invoice with header price dif
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_3_1_19 LIV Local Vendor DC and DSD Invoice with header price dif"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

''--------------------------------MIRO-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
call ClickButtonIfExist("Continue   \(Enter\)",True)

 Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetComboByKey("RM08M-VORGANG",DT_MIRO_6000_TRANSACTION)
Call PressEnter()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE_OCC1,False)
Call PressEnter()
Call TakeScreenShot()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call TakeScreenShot()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP_OCC1)
Call SetComboByKey("RM08M-XWARE_BNK","1")

Call SetTextboxNoLabel("RM08M-EBELN",0,DT_MIRO_6211_RM08MEBELN,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call PressEnter()
Call TakeScreenShot()


Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot()
Call GetStatusBar("item1","DT_DOC_NUM_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_DOC_NUM_OUTPUT&" created" )


Call LogOff()
Call FinalStatus ()


