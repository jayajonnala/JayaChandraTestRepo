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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Dunning Letters_p7_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 6th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage Dunning Letters_p7_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Dunning Letters_p7_TASE.xls"
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
'
'''----------------------Tcode FB01----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FB01_0100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_FB01_0100_POSTING_DATE),"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB01_0100_REFERENCE,False)
'Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB01_0100_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Value date","BSEG-VALUT","",Replace((DT_FB01_0300_VALUE_DATE),"/","."),False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0300_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT_OCC1,False)
'Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Document;Simulate")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Goto;First item")
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR,False)
Call SelectMenuBar("Goto;Next Item")
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR,False)
Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call SelectMenuBar("Document;Simulate General Ledger")
Call VerifyGridCellContent("",1,"Posting Key","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key","",DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call ClickButton("Back   \(F3\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_FB01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

