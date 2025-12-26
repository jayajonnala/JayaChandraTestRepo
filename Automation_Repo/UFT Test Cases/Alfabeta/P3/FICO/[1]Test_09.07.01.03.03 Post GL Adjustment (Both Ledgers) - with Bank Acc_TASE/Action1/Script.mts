

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.03.03 Post GL Adjustment (Both Ledgers) - with Bank Acc
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If

gstrTestCaseName = "Test_09.07.01.03.03 Post GL Ad"'justment (Both Ledgers) - with Bank Acc"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode FB01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 

Call TakeScreenShot()
 
Call SetTextbox("Period","BKPF-MONAT","",DT_FB01_0100_PERIOD,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_FB01_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_FB01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB01_0100_CURRENCYRATE,False)
Call TakeScreenShot()

Call FocusTextBox("Account","RF05A-NEWKO",False)
Call PressEnter()
Wait(2)
'Call PressEnter()
'Wait(2)	
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FB01_0300_TAX_CODE,False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_FB01_1007_COST_CENTER,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_FB01_1007_BUSINESS_AREA,False)
Call TakeScreenShot()

Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB01_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB01_0300_ACCOUNT,False)
Call FocusTextBox("Cost Center","COBL-KOSTL",False)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB01_0300_AMOUNT_OCC1,False)
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

Call SetTextbox("House Bank","BSEG-HBKID","",DT_FB01_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_FB01_0330_BSEGHKTID,False)
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyTextBoxContent("C", "RF05A-AZSAL",0,DT_FB01_0700_CHECK_TEXT_OF_C, False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()

If UCASE(GetStatusBar("item1","DT_DOC_1_OUTPUT")) <> UCASE(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR) Then
Call PressEnter()	
End If

Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_FB01_0100_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()




