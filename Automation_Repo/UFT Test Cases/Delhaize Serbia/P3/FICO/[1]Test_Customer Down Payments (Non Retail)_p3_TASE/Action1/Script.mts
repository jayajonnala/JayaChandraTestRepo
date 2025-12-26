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
'reload DS to update dates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Customer Down Payments (Non Retail)_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Down Payments (Non Retail)_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer Down Payments (Non Retail)_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
'Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call EndDateof445PeriodByDate(DT_TODAY,"DT_ENDING_DATE_PERIOD")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode FB05----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'Call SelectRadioButton("RF05A-XPOS1",DT_FB05_0122_TRANSFER_POSTING_WITH_CLEARING,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FB05_0122_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_FB05_0122_POSTING_DATE),"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB05_0122_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)
'Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB05_0122_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB05_0122_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
wait(1)

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("RF05A-XMWST","1",DT_FB05_0304_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0304_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB05_0304_TAX_CODE,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB05_0304_BUS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Display Long Texts for Item",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("1st line","EENO_DYNP-ZEILE","",DT_FB05_1001_1ST_LINE,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButtonIfExist("Copy text   \(F5\)",True)

'Capture the screenshot
Call TakeScreenShot()
Call SelectCheckbox("RF05A-XMWST","1",DT_FB05_0304_CALCULATE_TAX_OCC1,False)

Call ClickButton("Choose open items   \(F6\)",False)

Call SetTextbox("Account","RF05A-AGKON","",DT_FB05_0710_ACCOUNT,False)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_FB05_0710_ACCOUNT_TYPE,False)
Call SetTextbox("Special G/L ind","RF05A-AGUMS","",DT_FB05_0710_SPECIAL_GL_IND,False)
Call SelectCheckbox("RF05A-XNOPS","1",DT_FB05_0710_NORMAL_OI,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
wait(1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
'veryfy sattus bar content
Call GetStatusBar("item1","DT_FB05_3100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB05_3100_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

'select and activate respective doc and validation___________________________________________
Call ClickButton("Select All",False) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Deactivate Items",False) 
'Capture the screenshot
Call TakeScreenShot()

''''selecting cell of row:1 , Column "RSD Gross"
'''SAPGuiSession("transaction:=FB05").SAPGuiWindow("transaction:=FB05").SAPGuiTable("guicomponenttype:=80","name:=SAPDF05XTC_6102").SelectCell 1,"RSD Gross"
'''Call DoubleClick()
''''Capture the screenshot
'''Call TakeScreenShot()



Call SelectCellGuiTable("SAPDF05XTC_6102","Document Number","Document Number",DT_DOCUMENT,False)
Call ClickButtonIfExist("Activate Items",False)



'_________________________________________________________________________________________
Call ClickButton("Post   \(Ctrl\+S\)",False)
wait 2
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(2)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")
'veryfy sattus bar content
Call GetStatusBar("item1","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_FB05_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB05_0122_CHECK_TEXT_OF_STATUSBAR_OCC1)


Call SelectMenuBar("Document;Display")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Special G/L ind.","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("",1,"Amount","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",1,"Tax Code","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)

Call VerifyGridCellContent("",2,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",2,"Special G/L ind.","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
Call VerifyGridCellContent("",2,"Amount","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",2,"Tax Code","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)

Call VerifyGridCellContent("",3,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("",3,"Amount","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("",3,"Tax Code","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)

Call VerifyGridCellContent("",4,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("",4,"Amount","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)
Call VerifyGridCellContent("",4,"Tax Code","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MWSKZ)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

