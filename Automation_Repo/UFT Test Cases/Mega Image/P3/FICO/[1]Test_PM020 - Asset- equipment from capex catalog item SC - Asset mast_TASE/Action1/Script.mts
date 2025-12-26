
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PM020 - Asset- equipment from capex catalog item SC - Asset mast
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

gstrTestCaseName = "Test_PM020 - Asset- equipment from capex catalog item SC - Asset mast_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"
'
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-ASO3 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Asset", "ANLA-ANLN1", 0, DT_AS03_0100_ASSET, False)
Call SetTextbox("Sub-number","ANLA-ANLN2", 0, DT_AS03_0100_SUBNUMBER, False)
Call SetTextbox("Company Code","ANLA-BUKRS", 0, DT_AS03_0100_COMPANY_CODE, False)
Call TakeScreenSHot()

Call PressEnter()
Call TakeScreenSHot()

Call VerifyTextBoxContent("Description","ANLA-TXT50", 0, Lcase(DT_AS03_1140_CHECK_TEXT_OF_DESCRIPTION), False)
Call TakeScreenSHot()

Call SelectTab("TABSTRIP100", "Time-dependent", FAlse)
Call TakeScreenSHot()
Call VerifyTextBoxContent("Business Area","ANLZ-GSBER", 0, Lcase(DT_AS03_1145_CHECK_TEXT_OF_BUSINESS_AREA), False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL", 0, Lcase(DT_AS03_1145_CHECK_TEXT_OF_COST_CENTER), False)
Call TakeScreenSHot()

Call SelectTab("TABSTRIP100", "Allocations", FAlse)
Call TakeScreenSHot()
Call VerifyTextBoxContent("Evaluation group 2","ANLA-ORD42", 0, Lcase(DT_AS03_1160_CHECK_TEXT_OF_EVALUATION_GROUP_2), False)

Call SelectTab("TABSTRIP100", "Origin", FAlse)
Call TakeScreenSHot()
'Call VerifyTextBoxContent("Type name","ANLA-TYPBZ", 0, Lcase(DT_AS03_1181_CHECK_TEXT_OF_TYPE_NAME), False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR", 0, Lcase(DT_AS03_1182_CHECK_TEXT_OF_WBS_ELEMENT), False)

Call LogOff()
Call FinalStatus()
