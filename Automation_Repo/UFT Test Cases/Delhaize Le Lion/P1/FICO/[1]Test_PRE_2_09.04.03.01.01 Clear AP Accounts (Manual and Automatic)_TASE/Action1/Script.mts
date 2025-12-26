

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_2_09.04.03.01.01 Clear AP Accounts (Manual and Automatic)
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

gstrTestCaseName = "Test_PRE_2_09.04.03.01.01 Clear AP Accounts (Manual and Automatic)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_PRE_2_09.04.03.01.01 Clear AP Accounts (Manual and Automatic).xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'''
'''''''''--------TransactionCode-FB60----------''''
'''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F02_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F02_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0100_PSTKY,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F02_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F02_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F02_0100_REFERENCE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F02_0100_TYPE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()


Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F02_0302_PSTKY,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0302_TEXT,False)
Call SetTextbox("G/L Acc","BSEG-HKONT","",DT_GL_ACC,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0302_AMOUNT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F02_0302_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call SetTextbox("House bank","BSEG-HBKID","",DT_F02_0332_HOUSE_BANK,False)
Call PressEnter()

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F02_0300_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F02_0300_TEXT,False)
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot

Call SetTextbox("House Bank","BSEG-HBKID","",DT_F02_0330_HOUSE_BANK_OCC1,False)
Call SetTextbox("/","BSEG-HKTID","",DT_F02_0330_BSEGHKTID,False)
Call TakeScreenShot
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call TakeScreenShot
Call Logoff
Call FinalStatus()
