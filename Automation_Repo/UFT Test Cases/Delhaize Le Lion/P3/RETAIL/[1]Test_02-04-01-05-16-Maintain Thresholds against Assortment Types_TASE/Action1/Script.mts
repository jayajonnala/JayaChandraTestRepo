'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-16-Maintain Thresholds against Assortment Types  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-05-16-Main"
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

''''''--------TransactionCode-ZMDAS_ASORT_CHECK----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Assortment Type","S_ASTYP-LOW","",DT_ZMDAS_ASORT_CHECK_1000_ASSORTMENT_TYPE,False)
Call SetTextbox("Distribution Channel","P_VTWEG","",DT_ZMDAS_ASORT_CHECK_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Sales Organization","P_VKORG","",DT_ZMDAS_ASORT_CHECK_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Assortment","S_ASORT-LOW","",DT_ZMDAS_ASORT_CHECK_1000_ASSORTMENT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell",1,"Assortment","",DT_ZMDAS_ASORT_CHECK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASORT)
Call VerifyGridCellContentbyName("shell",1,"Assortment Type","",DT_ZMDAS_ASORT_CHECK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASSORTYP)
Call VerifyGridCellContentbyName("shell",1,"Number of items in layout module","",DT_ZMDAS_ASORT_CHECK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NR_OF_ITEMS)
Call VerifyGridCellContentbyName("shell",1,"Threshold value of assortment","",DT_ZMDAS_ASORT_CHECK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_THRESH_VAL)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
