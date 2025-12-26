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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Reverse AP Document_p13_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 15th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Reverse AR Document_p13_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\FICO\TASE_DT_Reverse AR Document_p13.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB02_0100_FISCAL_YEAR,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB02_0100_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB02_0100_DOCUMENT_NUMBER,False)
Call TakeScreenShot()
Call PressEnter()  
WAIT 5	
Call TakeScreenShot()
Call DoubleClickGuiGridCell("","",1,DT_FB02_0750_GRIDCELL_0_DESCRIPTION,False)
Wait(2)
Call TakeScreenShot()
Call SetTextbox("Invoice Ref\.","BSEG-REBZG","",DT_FB02_0301_INVOICE_REF,False)
Call PressEnter() 
wait 5
Call PressEnter() 
Call ClickButton("Save   \(Ctrl\+S\)",False)
Wait(2)
Call TakeScreenShot()
Call VerifyStatusBar(DT_FB02_0100_CHECK_TEXT_OF_STATUSBAR)
Call SetTcode(DT_FB02_0100_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB02_0100_FISCAL_YEAR_OCC1,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB02_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB02_0100_DOCUMENT_NUMBER_OCC1,False)
Call TakeScreenShot()
Call PressEnter()     ' 
Call TakeScreenShot()
Call DoubleClickGuiGridCell("","",1,DT_FB02_0750_GRIDCELL_0_DESCRIPTION_OCC1,False)
Wait(2)
Call TakeScreenShot()
Call VerifyTextBoxContent("Credit memo","BSEG-REBZG","",DT_FB02_0301_CHECK_TEXT_OF_CREDIT_MEMO,False)


Call LogOff()
Call FinalStatus ()



