

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.07.01 Manage Asset Trans Intra-Opco_Trans within CC_P1
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
'.................Test Script Name : Test_09.11.01.07.01 Manage Asset Trans Intra-Opco_Trans within CC_P1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.11.01.07.01 Manage Asset Trans Intra-Opco_Trans within CC_P1"
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

'----------------------Tcode ABUMN----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter Company Code
Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABUMN_0300_COMPANY_CODE,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABUMN_0300_ASSET,False)
Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABUMN_0206_TEXT,False)
Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_ABUMN_0320_EXISTING_ASSET,False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABUMN_0202_ASSET_VALUE_DATE),False)
Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_ABUMN_0201_POSTING_DATE),False)
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABUMN_0200_DOCUMENT_DATE),False)
Call TakeScreenShot()


'Click on Simulate
Call ClickButton("Simulate   \(F9\)",False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item2","DT_ASSET_TRANSC_DOCNO_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. GR02 "&DT_ASSET_TRANSC_DOCNO_OUTPUT)

''----------------------Tcode AS03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ABUMN_0100_OKCD) 
Call PressEnter() 
'Call CheckTCodeScreen(DT_ABUMN_0100_OKCD)
'Call PressEnter() 

'Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABUMN_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABUMN_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABUMN_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 

Call VerifyTextBoxContent("Deactivation on","ANLA-DEAKT",0,ConvertDate(DT_ABUMN_1142_CHECK_TEXT_OF_DEACTIVATION_ON),False)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

