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

'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name      : Test_Manage AR Bill of Exchange_p1_TASE
'.................Author : TCS          :Bitan
'................ Creation Date         : 19th April
'.................Modified By           :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage AR Bill of Exchange_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AR Bill of Exchange_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
'DataRowSet=2

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode FB05----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_FB05_0122_TRANSFER_POSTING_WITH_CLEARING,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDateFormat(DT_FB05_0122_DOCUMENT_DATE),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB05_0122_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB05_0122_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB05_0122_SGL_IND,False)


Call TakeScreenShot()
Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_2320_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB05_2320_TEXT,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",ConvertDateFormat(DT_FB05_2320_DUE_ON),False)


Call TakeScreenShot()
'______________________________________________________________________________________________
Call SetTextbox("Bank Key","BSEC-BANKL","",DT_FB05_2320_BANK_KEY,False)
Call FocusTextBox("Bank Key","BSEC-BANKL",False)
Call TakeScreenShot()
Call SendKey("{F2}")
wait(1)
Call PressEnter() 
wait(1)
Call SendKey("{ENTER}")
wait(1)
Call SendKey("{DOWN}")

Call TakeScreenShot()
Call SendKey("{F2}")

Call SetTextbox("Requested line item","\*BSEG-BUZEI","",DT_FB05_0610_REQUESTED_LINE_ITEM,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
'______________________________________________________________________________________________

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_2320_PSTKY,False)
'______________________________________________________________________________________________
Call FocusTextBox("Account","RF05A-NEWKO",False)
Call TakeScreenShot()
Call SendKey("{F4}")
wait(5)
Call SetTextbox("G/L account","G_SELFLD_TAB-LOW","",DT_FB05_0220_GL_ACCOUNT,True)
Call SetTextbox("G/L account","G_SELFLD_TAB-LOW","",DT_FB05_0220_GL_ACCOUNT,True)
Call TakeScreenShot()

Call PressEnter() 
Call ClickButton("Find   \(Ctrl\+F\)",True)

Call SetTextbox("Find","RSYSF-STRING","",DT_FB05_0800_FIND,True)
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)


Call TakeScreenShot()
Call ClickLabel(DT_FB05_0800_FIND,"",True)

Call TakeScreenShot()
Call ClickButton("Copy   \(Enter\)",True)
'______________________________________________________________________________________________
Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0300_AMOUNT,False)
Call TakeScreenShot()

Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus ()

