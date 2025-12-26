'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AR Bill of Exchange_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage AR Bill of Exchange_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AR Bill of Exchange_p4_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'----------------------Tcode F-28----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDateFormat(DT_F28_0103_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F28_0103_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F28_0103_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDateFormat(DT_F28_0103_POSTING_DATE),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F28_0103_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F28_0103_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F28_0103_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F28_0103_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-KONTO","",DT_F28_0103_ACCOUNT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F28_0103_AMOUNT,False)
Call SetTextbox("Value date","BSEG-VALUT","",ConvertDateFormat(DT_F28_0103_VALUE_DATE),False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F28_0103_ACCOUNT_OCC1,False)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_F28_0103_ACCOUNT_TYPE,False)
Call SetTextbox("Special G/L ind","RF05A-AGUMS","",DT_F28_0103_SPECIAL_GL_IND,False)

Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Wait(1)
Call TakeScreenShot()

Call ClickButton("Select All",False) 
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False) 
Call TakeScreenShot()

Call ClickButton("Field content search",False) 
Call SelectRadioButton("RF05A-XPOS1",DT_F28_2000_DOCUMENT_NUMBER,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call SetTextbox("From","RF05A-SEL01","",DT_F28_0731_FROM,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call TakeScreenShot()


Call TakeScreenShot()
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_F28_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)  'Value:=84.000,00

Call SelectCellGuiTable("SAPDF05XTC_6102","RSD Gross","Document Number",DT_F28_0731_FROM,False)
Call ClickButton("Activate Items",False) 
Call TakeScreenShot()
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_F28_6102_CHECK_TEXT_OF_NOT_ASSIGNED_OCC1,False)  'Value:=0,00

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_F28_0103_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F28_0103_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()

