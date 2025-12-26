

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P1
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
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode F-92----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F92_0100_DOCUMENT_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F92_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F92_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F92_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F92_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F92_0100_ACCOUNT,False)
Call TakeScreenShot()
Call PressEnter() 


'Enter details
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F92_0301_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_F92_0301_TAX_AMOUNT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F92_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F92_0301_ACCOUNT,False)
Call TakeScreenShot()
Call PressEnter() 


Call SelectCheckbox("RF05A-XAABG",0,DT_F92_0300_ASST_RETIREMENT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F92_0300_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F92_0300_TAX_CODE,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F92_0300_TEXT,False)
Call SetTextbox("Business Area","COBL-GSBER","",DT_F92_1007_BUSINESS_AREA,False)
Call SetTextbox("Trdg part\.BA","COBL-PARGB",""," ",False)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F92_1007_COST_CENTER,False)
Call TakeScreenShot()
Call PressEnter() 


Call SelectCheckbox("ANBZ-XVABG",0,DT_F92_0300_ASST_RETIREMENT,True)
Call SetTextbox("Asset","RA01B-ANLN1","",DT_F92_0210_ASSET,True)
Call SetTextbox("Asset Value Date","ANBZ-BZDAT","",ConvertDate(DT_F92_0210_ASSET_VALUE_DATE),True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)


Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'''Call VerifyTextBoxContent("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1",0,DT_F92_0700_CHECK_TEXT_OF_PK__BUSA_ACCT_EUR_AMOUNT_TAX_AMNT,False)

'Call GetTextboxValue("RF05A-AZEI1",1,"DT_F92_0700_CHECK_TEXT_OF_PK__BUSA_ACCT_EUR_AMOUNT_TAX_AMNT_OCC1",False)
'Call VerifyTextBoxContent("PK  BusA Acct                               EUR   Amount        Tax amnt","RF05A-AZEI1",1,DT_F92_0700_CHECK_TEXT_OF_PK__BUSA_ACCT_EUR_AMOUNT_TAX_AMNT_OCC1,False)
Call TakeScreenShot()
Call VerifyTextBoxContent("C","RF05A-AZSAL",0,DT_F92_0700_CHECK_TEXT_OF_C,False)
Call PressEnter()


'Post the Document
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOCUMENT_NO_OUTPUT&" was posted in company code GR02")

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

