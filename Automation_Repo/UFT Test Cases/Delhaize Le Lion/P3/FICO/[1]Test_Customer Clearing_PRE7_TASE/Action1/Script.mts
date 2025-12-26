'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Clearing_PRE7 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Clearing_PRE7"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-FB01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 


Call ClickButtonIfExist("Post with Reference Document   \(Shift\+F9\)",False)
Call SetTextbox("Document Number","BKPF-BELNR","",DT_FB01_0104_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0104_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BKPF-GJAHR","",DT_FB01_0104_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0100_POSTING_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Choose   \(F2\)",False)
Call SetTextbox("Requested line item","\*BSEG-BUZEI","","1",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Continue Processing   \(Shift\+F1\)",True)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Back   \(F3\)",False)
Call PressEnter()
Call TakeScreenShot
'Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call ClickButtonIfExist("Choose   \(F2\)",False)
Call SetTextbox("Requested line item","\*BSEG-BUZEI","","2",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'Call ClickButtonIfExist("Continue Processing   \(Shift\+F1\)",True)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetTextboxValue("MESSTXT1", "","DT_OP_FB01_0010_CHECK_TEXT_OF_MESSTXT1", True) 
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyTextBoxContent("Information Message", "MESSTXT1", "",lcase(DT_FB01_0010_CHECK_TEXT_OF_MESSTXT1), True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
