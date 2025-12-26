'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_PRE_04_E10G07P02S01V01 Client Listing for BE10
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_04_E10G07P02S01V01 Client Listing for BE10"
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

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
wait 10
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
wait 10
Call ClickButton("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode--SM35---------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Created by:","D0100-CREATOR","",DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY,False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1, False)
Call ClickButton("Process session   \(F8\)",False)
Call SelectRadioButton("D0300-BATCH", "Background", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
wait 30
Call ClickButton("Analyze session   \(F2\)",False)
wait 70
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Analyze session   \(F2\)",False)
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_SM35_1005_TO__OCC1),False)
Call TakeScreenShot
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message",4, "", "","DT_OP_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_3", False)
GetRowNo =4
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
