'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G10P06S13V35 Execute_Depreciation_Run
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G10P06S13V35 Execute"
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-AW01N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AW01N_0201_COMPANY_CODE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AW01N_0201_ASSET,False)
Call PressEnter()     
Call TakeScreenShot
Call GetGridContentByTitle("Planned values IFRS APC, depreciation", "", "Change",5,"DT_OP_AW01N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE")
Call ClickButton("Back   \(F3\)",False)

'''''--------TransactionCode-AFAB----------''''

Call SetTcode(DT_AW01N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SelectRadioButton("P_CHECKP", "Unplanned posting run", False)
Call SelectCheckbox("P_TEST",0, "ON", False)
Call SetTextbox("Company Code","P_BUKRS","",DT_AW01N_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal year","P_GJAHR","",DT_AW01N_1000_FISCAL_YEAR,False)
Call SetTextbox("Posting Period","P_BUPER","",DT_AW01N_1000_POSTING_PERIOD,False)
Call SetTextbox("Main asset number","S_ANLN1-LOW","",DT_AW01N_1000_MAIN_ASSET_NUMBER,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call ClickButtonIfExist("Yes", True)
Call TakeScreenShot
Call PressEnter()
GetRowNo = 4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'Call CheckifGuiLabelExists(DT_AW01N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AFAP)
Call VerifyGridCellContentbyName("shell",1, "Planned Amt for Dep./Interest", "", DT_AW01N_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_JENDE)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
