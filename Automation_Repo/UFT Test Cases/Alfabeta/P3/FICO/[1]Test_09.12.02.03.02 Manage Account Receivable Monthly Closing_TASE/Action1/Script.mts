		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.12.02.03.02 Manage Account Receivable Monthly Closing



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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_09.12.02.03.02 Manage Account Receivable Monthly Closing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call ClickButtonifExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Position\.\.\.",False)
Call TakeScreenShot
Call SetTextbox("OpCo Level","SVALD-VALUE", "", DT_ZFIGL_OPCO_LVL_0300_OPCO_LEVEL, True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call VerifyTableCellContent("3", "OpCo Level", "SAPLZFIGL_TAB_MAINTTCTRL_ZFIGL_OPCO_LVL", DT_ZFIGL_OPCO_LVL_0001_CHECK_TEXT_OF_TABLECELL_OPCO_LEVEL_0)
Call VerifyTableCellContent("3", "Description", "SAPLZFIGL_TAB_MAINTTCTRL_ZFIGL_OPCO_LVL", Lcase(DT_ZFIGL_OPCO_LVL_0001_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_0))

Call LogOff
Call FinalStatus()
