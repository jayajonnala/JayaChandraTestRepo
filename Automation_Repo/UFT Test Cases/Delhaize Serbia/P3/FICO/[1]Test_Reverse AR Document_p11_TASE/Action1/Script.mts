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
'.................Test Script Name : Test_Reverse AP Document_p11_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 15th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Reverse AR Document_p11_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\FICO\TASE_DT_Reverse AR Document_p11.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'' 

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_F80_1000_COMPANY_CODE,False)
Call SetTextbox("Reason for reversal","STOGRD","",DT_F80_1000_REASON_FOR_REVERSAL,False)

Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_YEAR,False)
Call TakeScreenShot()
Call ClickButton("%_BR_BELNR_%_APP_%-VALU_PUSH",False)

Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_F80_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_F80_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",False)
Call SelectCheckBox("TESTLAUF",0,"ON",False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_F80_1000_FISCAL_YEAR,False)
Call TakeScreenShot()

Call SelectRadioButton("BVOR_NO","Do not process",False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call CheckifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_NO_NAME)
Call CheckifGuiLabelExists(DT_F80_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call ClickButton("Back   \(F3\)",False)
Call SelectCheckBox("TESTLAUF",0,"OFF",False)
Call ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar(DT_F80_0120_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

Call GetLabelContentByRefLabel(DT_F80_3010_TABLECELL_SINGLE_VALUE_0,-112,-16,"DT_REVERSED_WITH_DOCUMENT_1600000209_OUTPUT",False)
Call GetLabelContentByRefLabel(DT_F80_3010_TABLECELL_SINGLE_VALUE_0_OCC1,-112,-16,"DT_REVERSED_WITH_DOCUMENT_1600000210_OUTPUT",False)

Call LogOff()
Call FinalStatus ()




