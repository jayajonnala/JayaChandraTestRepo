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
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Customer Down Payments (Non Retail)_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Down Payments (Non Retail)_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Customer Down Payments (Non Retail)_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_DOCHEADER_TEXT",(Cint(DT_INCREMENT_DOCHEADER_TEXT)+1))
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",(Cint(DT_INCREMENT_REFERENCE)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_DOCHEADER_TEXT",(Cint(DT_INCREMENT_DOCHEADER_TEXT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''----------------------Tcode F-29----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

''Call SetTextbox("Document Date","BKPF-BLDAT","",DT_F29_0111_DOCUMENT_DATE,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F29_0111_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F29_0111_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F29_0111_COMPANY_CODE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F29_0111_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F29_0111_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F29_0111_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F29_0111_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F29_0111_ACCOUNT,False)
Call SetTextbox("Special G/L ind","RF05A-UMSKZ","",DT_F29_0111_SPECIAL_GL_IND,False)
Call SetTextbox("Account","RF05A-KONTO","",DT_F29_0111_ACCOUNT_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F29_0111_AMOUNT,False)
''Call SetTextbox("Value date","BSEG-VALUT","",DT_F29_0111_VALUE_DATE,False)
Call SetTextbox("Value date","BSEG-VALUT","",Replace(DT_F29_0111_VALUE_DATE,"/","."),False)
'Call SelectCheckbox("NAME","",DT_F29_1702_SAPMF05ATC_1702,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Requests   \(F6\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call GetStatusBar("item1","DT_F29_1702_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F29_1702_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
''
'''selecting cell of row:1 , Column "Document Number"
''SAPGuiSession("transaction:=F-29").SAPGuiWindow("transaction:=F-29").SAPGuiTable("guicomponenttype:=80","name:=SAPMF05ATC_1702").SelectCell 1,"Document Number"
''Call DoubleClick()
'''Capture the screenshot
''Call TakeScreenShot()

Call SelectRowGuiTable("SAPMF05ATC_1702","Document Number",DT_DOCUMENT,False)
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)

'Capture the screenshot
Call TakeScreenShot()
wait 3
Call PressEnter
Call GetStatusBar("item1","DT_F29_0111_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_F29_0111_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F29_0111_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

Call SelectMenuBar("Document;Display")
Wait(2)
CAll SelectRowGuiGridbyRowNo("","",DT_F29_0750_GRIDCELL_1_PK,False)
'Capture the screenshot
Call TakeScreenShot()
Call GetGridContentByTitle("","","Posting Key",1,"DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OUTPUT")
Call GetGridContentByTitle("","","Account",1,"DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OUTPUT")
Call GetGridContentByTitle("","","Amount",1,"DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET_OUTPUT")
Call GetGridContentByTitle("","","Posting Key",2,"DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OUTPUT")
Call GetGridContentByTitle("","","Account",2,"DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OUTPUT")
Call GetGridContentByTitle("","","Amount",2,"DT_F29_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()





