'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AP Bill of Exchange_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 11th March
'.................Modified By :
'.................Modified Date/Details :
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage AP Bill of Exchange_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AP Bill of Exchange_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'''----------------------Tcode FB05----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RF05A-XPOS1",DT_FB05_0122_TRANSFER_POSTING_WITH_CLEARING,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB05_0122_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB05_0122_SGL_IND,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FB05_0122_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Choose open items   \(F6\)",False)
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_FB05_2320_DUE_ON),"/","."),False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_2320_AMOUNT,False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Special G/L ind","RF05A-AGUMS","",DT_FB05_0710_SPECIAL_GL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Select All",False) 
Call ClickButton("Deactivate Items",False) 
Call SelectCellGuiTable("SAPDF05XTC_6102","Document Number","Document Number",DT_DOC_NUMBER,False)
Call TakeScreenShot()
Call ClickButton("Sort in Descending Order",False) 
Call TakeScreenShot()
Call ClickButton("Field content search",False) 
Call SelectRadioButton("RF05A-XPOS1","Document Number",True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call SetTextbox("From","RF05A-SEL01","",DT_DOC_NUMBER,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call TakeScreenShot()
Call ClickButton("Select All",False)
Call TakeScreenShot()
Call SelectCellGuiTable("SAPDF05XTC_6102","Document Number","Document Number",DT_DOC_NUMBER,False)
Call TakeScreenShot()
Call ClickButton("Sort in Ascending Order",False) 
Call ClickButton("Activate Items",False) 
Call TakeScreenShot()
''''''''''''''''''''''
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_FB05_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)  'Value:=0,00

Call SelectMenuBar("Document;Simulate")
'Capture the screenshot
Call TakeScreenShot()
''DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT  verification
''if script fails with bellow line where the value is hard coaded,,, any of the next three lines to be "un commented" or have to change the hardcoded value
Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,"001 39X      0010030933 STANČIĆ DOO               100.000,00-",False)
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT,False)
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,UCase(DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT),False)
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,LCase(DT_FB05_0700_CHECK_TEXT_OF_PK_BUSA_ACCT),False)


Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call PressEnter()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR)

Call SelectMenuBar("Document;Display")
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'validate grid components
call VerifyGridCellContent("",1,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Special G/L ind.","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call TakeScreenShot()

Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

