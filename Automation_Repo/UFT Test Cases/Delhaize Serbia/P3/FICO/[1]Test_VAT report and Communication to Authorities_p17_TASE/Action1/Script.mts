

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p17
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p17"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''''''--------TransactionCode-S_ALR_87012357----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)", False)

Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RFUMSV00.*","",3, True)
Call ClickButton("Choose   \(F2\)", True)
Call SelectCheckbox("PAR_DEF",0,"OFF",False)
Call SetTextbox("Layout","PAR_VAR1","",DT_LAYOUT1,False)
Call SetTextbox("Layout","PAR_VAR3","",DT_LAYOUT2,False)
Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_S_ALR_87012357_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_S_ALR_87012357_1000_COMPANY_CODE,False)
''Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","","2021",False)
Call SetTextboxNoLabel("BR_GJAHR-LOW","",DT_YEAR,False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012357_1000_TO),False)
'Call SetTextbox("Tax on sales/purchases code","SELC_MWKZ-LOW","","",False)
Call SetTextboxNoLabel("SEL_MWKZ-LOW","","",False)
Call SelectCheckbox("PAR_NOHE",0,"OFF",False)
Call SelectCheckbox("PAR_LIS2",0,"ON",False)
Call SelectCheckbox("PAR_LIS1",0,"ON",False)
Call SetTextbox("Program run date","PAR_LAUD","",ConvertDate(DT_S_ALR_87012357_1000_PROGRAM_RUN_DATE),False)
Call SetTextbox("Identification","PAR_LAUI","",DT_S_ALR_87012357_1000_IDENTIFICATION,False)
Call TakeScreenShot
Call PressEnter()
''Call SetTextboxNoLabel("SEL_MWKZ-LOW","","",False)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_RS01,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_MT,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC1,0)

''''''--------TransactionCode-/CCEE/RSFI_RFUVDE----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE2)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Get Variant...   \(Shift\+F5\)", False)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program /CCEE/YUFI_RFUVDE00","",1, True)
Call TakeScreenShot
Call ClickButton("Choose   \(F2\)", True)
Call SetTextbox("Additional Run ID","PAR_LAUI","",DT_CCEERSFI_RFUVDE_1000_ADDITIONAL_RUN_ID,False)
Call SetTextbox("Run Date","PAR_LAUD","",ConvertDate(DT_CCEERSFI_RFUVDE_1000_RUN_DATE),False)
Call SetTextbox("Company Code","SEL_BKRS-LOW","",DT_CCEERSFI_RFUVDE_1000_COMPANY_CODE,False)
Call ClickButton("Execute   \(F8\)", False)
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextbox("Output Device","SSFPP-TDDEST","",DT_CCEERSFI_RFUVDE_0100_OUTPUT_DEVICE,True)
Call ClickButton("Print preview   \(F8\)", True)
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
