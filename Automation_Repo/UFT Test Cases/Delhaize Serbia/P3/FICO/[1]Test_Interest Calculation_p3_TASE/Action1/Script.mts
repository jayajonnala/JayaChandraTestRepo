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

'....................Mandatory Initial Call only in First Component in a Test Scenario
''Reload DataSheet to updates and calculations
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Interest Calculation_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 5th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Interest Calculation_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Interest Calculation_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '..
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
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
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB01_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FB01_0100_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call PressEnter()  
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0303_AMOUNT,False)
Call PressEnter()  
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB01_0303_BUS_AREA,False)
Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_FB01_0303_DUE_ON),"/","."),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0303_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0303_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call PressEnter() 

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)

Call PressEnter()
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectMenuBar("Document;Simulate")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Company Code","RF05A-AZBUK","",DT_FB01_0750_CHECK_TEXT_OF_COMPANY_CODE,False)

'due to the following special char :"Đ", bellow verification is hardcoaded.
'Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,Ucase(DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL),False)
Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",0,"001 092 A001 0011338340 ANĐA TUR                      210,00",False)

Call VerifyTextBoxContent("PK  BusA Acct                               RSD   Amount        Tax amnt","RF05A-AZEI1",1,Lcase(DT_FB01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL),False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If invoice is generated
Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

