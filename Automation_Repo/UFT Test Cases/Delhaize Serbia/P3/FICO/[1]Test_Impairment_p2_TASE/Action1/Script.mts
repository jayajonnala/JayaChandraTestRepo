'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Impairment_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th March
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

gstrTestCaseName = "Test_Impairment_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F90_0100_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F90_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F90_0100_COMPANY_CODE,False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace(DT_F90_0100_POSTING_DATE,"/","."),False)
Call SetTextbox("Period","BKPF-MONAT","",DT_F90_0100_PERIOD,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F90_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F90_0100_REFERENCE,False)
Call SetTextbox("Doc.Header Text","BKPF-BKTXT","",DT_F90_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0100_ACCOUNT,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
'Call TakeScreenShot()
'Call PressEnter()  
Call SelectCheckbox("BKPF-XMWST",0,"OFF",False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0302_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","","",False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_F90_0302_BUS_AREA,False)
'Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_F90_0302_PAYT_TERMS,False)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_F90_0302_PAYT_TERMS,False)
Call SetTextbox("Bline Date","BSEG-ZFBDT","",Replace(DT_F90_0302_BLINE_DATE,"/","."),False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F90_0302_TTYPE,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  

Call TakeScreenShot()
Call PressEnter()  
'''Capture the screenshot
Call TakeScreenShot()
''Call PressEnter()  
''WAIT 5
''Call PressEnter()  
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0302_AMOUNT,False)


Call PressEnter()  

Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F90_0100_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()




