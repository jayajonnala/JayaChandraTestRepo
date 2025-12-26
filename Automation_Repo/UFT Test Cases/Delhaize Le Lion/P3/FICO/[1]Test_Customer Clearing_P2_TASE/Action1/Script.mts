'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Customer Clearing_PRE2
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Customer Clearing_P2"
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

''--------TransactionCode-SM35----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("From:","D0100-VON","",ConvertDate(DT_SM35_1005_FROM),False)
Call SetTextbox("Created by:","D0100-CREATOR","",DT_SM35_1005_CREATED_BY,False)
Call PressEnter()
Call TakeScreenShot
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Process session   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("D0300-BATCH", "Background", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
wait 20
Call SetTextbox("From:","D0100-VON","",ConvertDate(DT_SM35_1005_FROM),False)
Call SetTextbox("Created by:","D0100-CREATOR","",DT_SM35_1005_CREATED_BY,False)
Call PressEnter()
Call VerifyTableCellContent(1, "Status", "SAPMSBDC_CCTC_APQI","STA_OK")
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Analyze session   \(F2\)",False)
Call TakeScreenShot
Call SelectTab("TAB_DYNPRO"," Log created on "&ConvertDate(DT_SM35_1005_TO__OCC1),False)
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message", 4, "", "","DT_OP_GET_DOCUMENT", False)
GetRowNo =2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'Call VerifyTableCellContent(4, "Message", "RSBDC_ANALYSETC_PROTOCOL", DT_CHECK)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
