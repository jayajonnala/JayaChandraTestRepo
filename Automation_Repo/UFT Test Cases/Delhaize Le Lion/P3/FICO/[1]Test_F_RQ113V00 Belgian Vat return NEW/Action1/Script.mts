'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F_RQ113V00 Belgian Vat return NEW  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_F_RQ113V00 Belgian Vat return NEW"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
'Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-/nfotv----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Dominant Enterprise/Comp\. Code","SO_ORG-LOW","",DT_FOTV_1000_DOMENTCOCODE,False)
Call SetTextbox("Return Type","PA_TYP","",DT_FOTV_1000_RETURN_TYPE,False)
Call SetTextboxNoLabel("PA_CTRY","",DT_FOTV_1000_REPORTING_COUNTRY,False)
Call SetTextbox("ReportingPeriod","SO_PER-LOW","",DT_REPORTINGPERIOD,False)
Call SetTextbox("Rep\. Year","PA_RJAHR","",DT_REP_YEAR,False)
Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot    
Call ClickButton("Refresh   \(Ctrl\+F2\)", False)
Call ClickButton("Display Details on Package   \(Shift\+F7\)", False)
Call TakeScreenShot
Call ClickButton("Last Page   \(Ctrl\+Page down\)", False)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC1, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC2, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC3, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC4, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC5, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC6, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC7, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC8, 1)
Call VerifyifGuiLabelExists_ByIndex(DT_FOTV_0120_CHECK_TEXT_OF_NO_NAME_OCC9, 1)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
