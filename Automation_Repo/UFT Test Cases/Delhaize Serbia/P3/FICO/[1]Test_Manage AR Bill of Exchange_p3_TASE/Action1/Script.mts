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

''Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AR Bill of Exchange_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Manage AR Bill of Exchange_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AR Bill of Exchange_p3_TASE.xls"
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

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)




''----------------------Tcode FB05---------------------------

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
Call SetTextbox("Due on","BSEG-ZFBDT","",ConvertDateFormat(DT_FB05_2320_DUE_ON),False)
Call TakeScreenShot()

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
'Call ClickButtonifexist("Continue Processing   \(Shift\+F1\)",True)
Call SendKey("{ENTER}")
wait(1)
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call PressEnter() 
wait(1)
Call TakeScreenShot()

Call ClickButton("Select All",False) 
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False) 
Call TakeScreenShot()

Call SelectCellGuiTable("SAPDF05XTC_6102","Document Number","Document Number",DT_FB05_0731_FROM,False)
Call TakeScreenShot()

Call ClickButton("Field content search",False) 
Call SelectRadioButton("RF05A-XPOS1",DT_FB05_2000_DOCUMENT_NUMBER,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call SetTextbox("From","RF05A-SEL01","",DT_FB05_0731_FROM,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True) 
Call TakeScreenShot()


Call TakeScreenShot()
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_FB05_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)  'Value:=84.000,00

Call SelectCellGuiTable("SAPDF05XTC_6102","RSD Gross","RSD Gross",DT_FB05_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)
Call ClickButton("Activate Items",False) 

Call TakeScreenShot()
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB","",DT_FB05_6102_CHECK_TEXT_OF_NOT_ASSIGNED_OCC1,False)  'Value:=0,00

Call ClickButton("Post   \(Ctrl\+S\)",False)
' FocusTextBox(attachedText, textboxName, blnIsItPopup)
'FocusTextBox


Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()

