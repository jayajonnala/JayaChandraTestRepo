

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Pre_Asset_Post_Invoice_01
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
'.................Test Script Name : Test_Pre_Asset_Post_Invoice_01
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Pre_Asset_Post_Invoice_01"
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

'----------------------Tcode F-90----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter details
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_F90_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",ConvertDate(DT_F90_0100_POSTING_DATE),False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F90_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_F90_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0100_ACCOUNT,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F90_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F90_0100_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter() 

'Enter details
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_F90_0302_TAX_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F90_0302_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F90_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F90_0302_ACCOUNT,False)
Call SetTextbox("TType","RF05A-NEWBW","",DT_F90_0302_TTYPE,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F90_0305_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F90_0305_TAX_CODE,False)
Call SetTextbox("Quantity","BSEG-MENGE","",DT_F90_0305_QUANTITY,False)
Call TakeScreenShot()

'Click on Save Buton
Call SelectMenuBar("Document;Simulate")
Call TakeScreenShot()

Call  VerifyTextBoxContent("C","RF05A-AZSAL",0,DT_F90_0700_CHECK_TEXT_OF_C,False)

Call PressEnter() 

'Post the Document
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOCUMENT_NO_OUTPUT&" was posted in company code GR02")

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

