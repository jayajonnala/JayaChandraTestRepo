

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.05.01 Automated posting for exchanges difference
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.07.01.05.01 Auto post"'mated posting for exchanges difference"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''


'----------------------Tcode FB01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",False)
Call SetTextbox("Company Code","S_CCODE-LOW","",DT_FAGL_FC_VAL_1000_COMPANY_CODE,False)
Call SetTextbox("Valuation Key Date","P_VDATE","",ConvertDate(DT_FAGL_FC_VAL_1000_VALUATION_KEY_DATE),False)
Call SetTextbox("Valuation Area","P_VAREA","",DT_FAGL_FC_VAL_1000_VALUATION_AREA,False)
Call TakeScreenShot()
 
Call SetTextbox("Posting period","P_PPERID","",DT_FAGL_FC_VAL_1010_POSTING_PERIOD,False)
Call SetTextbox("Posting date","P_PPDATE","",ConvertDate(DT_FAGL_FC_VAL_1010_POSTING_DATE),False)
Call SetTextbox("Document date","P_PDDATE","",ConvertDate(DT_FAGL_FC_VAL_1010_DOCUMENT_DATE),False)
Call TakeScreenShot()

Call FocusTextBox("Posting period","P_PPERID",False)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB","Open Items: Subledger",False)
Call TakeScreenShot()

Call SelectCheckbox("P_APOI",0,DT_FAGL_FC_VAL_1030_VALUATE_VENDORS,False)
Call TakeScreenShot()

Call SelectCheckbox("P_AROI",0,DT_FAGL_FC_VAL_1030_VALUATE_CUSTOMERS,False)
Call SetTextbox("Fiscal Year","S_FYEAR-LOW","",DT_FAGL_FC_VAL_1030_FISCAL_YEAR,False)
Call SetTextbox("Document Number","S_DOCNR-LOW","",DT_FAGL_FC_VAL_1030_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

Call FocusTextBox("Fiscal Year","S_FYEAR-LOW",False)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB","Open Items: G/L Accounts",False)
Call TakeScreenShot()

Call SelectCheckbox("P_GLOI",0,DT_FAGL_FC_VAL_1020_VALUATE_GL_ACCOUNTS,False)
Call TakeScreenShot()

Call SelectTab("TABSTRIP_SCRN_TAB","Output / Technical Settings",False)
Call TakeScreenShot()

Call SelectCheckbox("P_ALTACT",0,DT_FAGL_FC_VAL_1090_ALTERNATIVE_ACCOUNT_NUMBER,False)
Call TakeScreenShot()

Call PressEnter()
Wait(2)
Call TakeScreenShot()


Call ClickButton("Execute   \(F8\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_FAGL_FC_VAL_0120_CHECK_TEXT_OF_NO_NAME)

Call ClickButton("2 Postings   \(Shift\+F6\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyifGuiLabelExistsByRelativeid(DT_FAGL_FC_VAL_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[8,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_FAGL_FC_VAL_0120_CHECK_TEXT_OF_11924012,"wnd\[0\]/usr/lbl\[70,9\]")

Call LogOff()
Call FinalStatus ()

