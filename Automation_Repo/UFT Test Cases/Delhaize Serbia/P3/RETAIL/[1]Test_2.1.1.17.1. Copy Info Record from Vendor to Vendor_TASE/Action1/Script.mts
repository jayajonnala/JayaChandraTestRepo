

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.17.1. Copy Info Record from Vendor to Vendor
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

gstrTestCaseName = "Test_2.1.1.17.1. Copy Info Record from Vendor to Vendor"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode ZMDPU_INFOREC_COPY----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call CheckScreen(DT_SAPTRANSACTIONCODE,Lcase(DT_ZMDPU_INFOREC_COPY_1000_CHECK_TEXT_OF_TITL))

'''Call SelectCheckbox("P_NOTDEL",0,DT_ZMDPU_INFOREC_COPY_1000_DO_NOT_COPY_INFORECORDS_MARKED_FOR_DELETION_,False)
Call SelectCheckbox("P_MRKORI",0,DT_ZMDPU_INFOREC_COPY_1000_MARK_ORIGINAL_INFORECORD_AS_DELETED_,False)
Call SelectCheckbox("P_COPINF",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_INFORECORD_AND_PO_TEXT_,False)
Call SelectCheckbox("P_COPSUB",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_SUBRANGE_ASSIGNMENT_,False)
Call SelectCheckbox("P_MRKGTI",0,DT_ZMDPU_INFOREC_COPY_1000_MARK_GTINS_AS_TARGET_VENDOR_SPECIFIC_,False)
Call SelectCheckbox("P_COPPRI",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_PRICING_CONDITIONS,False)
Call SelectCheckbox("P_MRKVEN",0,DT_ZMDPU_INFOREC_COPY_1000_MARK_NEW_VENDOR_AS_REGULAR_,False)
Call SelectCheckbox("P_SRCVEN",0,DT_ZMDPU_INFOREC_COPY_1000_USE_SOURCE_VENDOR_AS_PRIOR_VENDOR_,False)
Call SelectCheckbox("P_COPVEN",0,DT_ZMDPU_INFOREC_COPY_1000_COPY_VENDOR_ARTICLE_CODE_,False)

Call SetTextbox("Purch\. Organization","P_EKORG","",DT_ZMDPU_INFOREC_COPY_1000_PURCH_ORGANIZATION,False)
Call SetTextbox("Source Vendor","P_SRC_V","",DT_ZMDPU_INFOREC_COPY_1000_SOURCE_VENDOR,False)
Call SetTextbox("Target Vendor","P_TGT_V","",DT_ZMDPU_INFOREC_COPY_1000_TARGET_VENDOR,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDPU_INFOREC_COPY_1000_ARTICLE,False)

Call TakeScreenShot()

Call PressEnter() 

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call SelectRowGuiGrid("",0,"Purchasing Organization",DT_ZMDPU_INFOREC_COPY_1000_PURCH_ORGANIZATION,False)

Call TakeScreenShot()

Call ClickButton("PROCESS   \(F9\)",False)

Call TakeScreenShot()

Call ClickButton("YES",True)

Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_ZMDPU_INFOREC_COPY_0500_CHECK_TEXT_OF_STATUSBAR)

Call SelectRowGuiGridbyRowNo("",0,1,False)
Call DoubleClickGuiGridCell("",0,1,"Inforecord number",False)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call VerifyTextBoxContent("Info record","EINA-INFNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_INFO_RECORD,False)
Call VerifyTextBoxNoLabelContent("EINA-LIFNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_VENDOR,False)
'Call VerifyTextBoxContent("Vendor","EINA-LIFNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_VENDOR,False)
Call VerifyTextBoxContent("Article","EINA-MATNR",0,DT_ZMDPU_INFOREC_COPY_0101_CHECK_TEXT_OF_ARTICLE,False)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
