		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_2_09.06.02.01.03 Clear AR Accounts (Manual and Automatic)


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

gstrTestCaseName = "Test_PRE_2_09.06.02.01.03 Clear AR Accounts (Manual and Automatic)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-F-21----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_FB70_0100_DOCUMENT_DATE), False)
Call SetTextbox("Type","BKPF-BLART","",DT_FB70_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB70_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB70_0100_PSTKY,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FB70_0100_CURRENCYRATE,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB70_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB70_0300_AMOUNT,False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_FB70_1016_PROFIT_CTRS,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB70_0300_TEXT,False)
Call TakeScreenShot
Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)


Call SetTextbox("House Bank","BSEG-HBKID","",DT_FB70_0330_HOUSE_BANK,False)
Call SetTextbox("/","BSEG-HKTID","",DT_FB70_0330_BSEGHKTID,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FB70_0330_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FB70_0330_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FB70_0301_AMOUNT,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_FB70_0301_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FB70_0301_TEXT,False)
Call TakeScreenShot
Call PressEnter()

Call SelectMenuBar("Document;Simulate")

Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_FB70_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_FB70_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB70_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB70_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call LogOff'
Call FinalStatus()


