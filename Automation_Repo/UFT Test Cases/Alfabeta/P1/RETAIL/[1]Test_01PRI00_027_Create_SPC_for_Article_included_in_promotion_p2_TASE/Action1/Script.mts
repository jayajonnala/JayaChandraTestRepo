

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01PRI00_027_Create_SPC_for_Article_included_in_promotion_p2
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

gstrTestCaseName = "Test_01PRI00_027_Create_SPC_for_Article_included_in_promotion_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_027_Create_SPC_for_Article_included_in_promotion_p2_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
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

''--------------------------------------------  VKP7------------------------------------------------
Call SetTextbox("Pricing document","P_KBELN","",DT_VKP7_1000_PRICING_DOCUMENT,False)
Call ClickButton("Execute   \(F8\)",False)

Call VerifyGridCellContent("",1,"KBELN",0,DT_VKP7_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KBELN)
Call GetGridContent("",0,"MATNR",1,"<NA>","<NA>","DT_VKP7_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")
Call VerifyGridCellContent("",1,"BPSTA",0,DT_VKP7_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BPSTA)
Call VerifyGridCellContent("",1,"ENDPR",0,DT_VKP7_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ENDPR)



Call LogOff()
Call FinalStatus ()


