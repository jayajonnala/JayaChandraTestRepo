

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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


gstrTestCaseName = "Test_Cash Invoice Payment_C1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.12.01 Cash Invoice Payment_Movement Type F900_Retail_C1E_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------/POSDW/MON0 -----------------------------------

Call SelectTab("TABSTRIP_WB_TAB","Stores",False)
Call SetTextbox("Store","SO_STORE-LOW","",DT_POSDWMON1_0201_STORE,False) 
Call SetTextbox("Posting Date","SO_DATE-LOW","",ConvertDate(DT_POSDWMON1_0301_POSTING_DATE),False) 
Call TakeScreenSHot()
Call SelectTab("TABSTRIP_WB_TAB2","Header",False)
Call SetTextbox("Transaction Type","SO_TTCOD-LOW","",DT_POSDWMON1_0302_TRANSACTION_TYPE,False)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
call ActivateItemGuiTree(1,DT_ITEM_PATH,DT_ITEM_TEXT)
Call TakeScreenSHot()
call VerifyGridCellContent("",1,"Transaction Type",0,DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE)
call GetGridContent("",1,"Sales Totals Per Transaction",1,"Transaction Type",DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE,"DT_TurnOver")
Call TakeScreenSHot()
Call LogOff()
Call FinalStatus ()
