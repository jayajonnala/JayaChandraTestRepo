

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_Get_2_components_from_BOM
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

gstrTestCaseName = "Test_PRE_Get_2_components_from_BOMO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------ZMDAM_BOM_REPORT-----------------------------

Call SetTextbox("Purchasing group","S_WEKGR-LOW","",DT_ZMDAM_BOM_REPORT_1000_PURCHASING_GROUP,False)
Call SetTextbox("Link Type","S_STLAN-LOW","",DT_ZMDAM_BOM_REPORT_1000_LINK_TYPE,False)
Call SetTextbox("Header Article","S_MATNR-LOW","",DT_ZMDAM_BOM_REPORT_1000_HEADER_ARTICLE,False)
Call TakeScreenshot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot()

Call VerifyGridCellContent("",1,"MATNR",0,DT_ZMDAM_BOM_REPORT_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",2,"MATNR",0,DT_ZMDAM_BOM_REPORT_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MATNR)
Call GetGridContent("",0,"IDNRK",1,"<NA>","<NA>","DT_ZMDAM_BOM_REPORT_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_IDNRK_OUTPUT")
Call GetGridContent("",0,"IDNRK",2,"<NA>","<NA>","DT_ZMDAM_BOM_REPORT_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_IDNRK_OUTPUT")
Call GetGridContent("",0,"MENGE",1,"<NA>","<NA>","DT_ZMDAM_BOM_REPORT_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MENGE_OUTPUT")
Call GetGridContent("",0,"MENGE",2,"<NA>","<NA>","DT_ZMDAM_BOM_REPORT_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MENGE_OUTPUT")
'Call VerifyGridCellContent("",1,"Retention Level",0,DT_CHECK_SUM_EQUALS_TO_1)

Call LogOff()
Call FinalStatus ()


