

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_VAT report and Communication to Authorities_p13
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

gstrTestCaseName = "Test_VAT report and Communication to Authorities_p13"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
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
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTextbox("Created By","ENAME-LOW","","",True)
'Call SetTextbox("Variant","V-LOW","",DT_S_ALR_87012357_0600_GRIDCELL_42_VARIANT_NAME,True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Variant Catalog for Program RFUMSV00.*","",1, True)
Call ClickButton("Choose   \(F2\)", True)
Call SelectCheckbox("PAR_DEF",0,"OFF",False)
Call SetTextbox("Layout","PAR_VAR1","",DT_LAYOUT1,False)
''Call SetTextbox("Layout","PAR_VAR2","",DT_LAYOUT2,False)
''Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_S_ALR_87012357_1000_DOCUMENT_NUMBER,False)
''Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",Year(Date),False)
Call SetTextboxNoLabel("BR_BELNR-LOW","",DT_S_ALR_87012357_1000_DOCUMENT_NUMBER,False)
Call SetTextboxNoLabel("BR_GJAHR-LOW","",Year(Date),False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012357_1000_POSTING_DATE),False)
Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012357_1000_TO),False)
'''''Call SetTextbox("Tax on sales/purchases code","SELC_MWKZ-LOW","","",False)

Call SetTextboxNoLabel("SEL_MWKZ-LOW","","",False)

'Call SelectCheckbox("PAR_NOHE",0,"OFF",False)
Call SetTextbox("Program run date","PAR_LAUD","",ConvertDate(DT_S_ALR_87012357_1000_PROGRAM_RUN_DATE),False)
Call SetTextbox("Identification","PAR_LAUI","",DT_S_ALR_87012357_1000_IDENTIFICATION,False)
'Call SetTextbox("Fiscal period","SEL_MONA-LOW","","",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot
Call SetVerticalScrollBar(2,False)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_TA,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_RS01,0)

Call SetHorizontalScrollBar(150,False)

Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME,0)

Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC2,0)
Call VerifyifGuiLabelExists_ByIndex(DT_S_ALR_87012357_0120_CHECK_TEXT_OF_NO_NAME_OCC3,1)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()



''*********************************************End Of Script*********************************************************************

