'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage AP Guarantees_p1_TASE 
'.................Author : TCS        :Bitan
'................ Creation Date    : 09th March
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

gstrTestCaseName = "Test_Manage AP Guarantees_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage AP Guarantees_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_REF_INC",(Cint(DT_REF_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''----------------------Tcode FB05----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
'Replace((DT_DATE),"/",".")
Call TakeScreenShot()
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FB05_0122_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB05_0122_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB05_0122_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_FB05_0122_POSTING_DATE),"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FB05_0122_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB05_0122_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FB05_0122_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FB05_0122_DOCHEADER_TEXT,False)
Call SelectRadioButton("RF05A-XPOS1","Transfer posting with clearing",False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0122_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0122_ACCOUNT,False)

'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0300_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB05_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB05_0300_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_FB05_0302_GL_ACC,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB05_0302_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB05_0302_BUS_AREA,False)
'Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_FB05_0302_PAYT_TERMS,False)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_FB05_0302_PAYT_TERMS,False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",Replace((DT_FB05_0302_BLINE_DATE),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()
Call SelectMenuBar("Document;Simulate")
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()
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
Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
Wait(1)
'validate grid components
call VerifyGridCellContent("",1,"Company code","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
call VerifyGridCellContent("",1,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",1,"Account","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Posting Key","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",2,"Account","",DT_FB05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


