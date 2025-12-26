
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Creation Goods receipts in FIORI_TASE
'.................Author : TCS  
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//--------------------------------------------------------------------' ActivateSAPUITreeFrame(creationTime, frameTitle, htmlTag, htmlId, innertext, className, index, title, path)

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


gstrTestCaseName = "TC_07_Test_Production order confirmation in FIORI_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
''''Close All Browser
Call CloseAllBrowsers()
''launch and Login SRM Application
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(10)
Call CaptureWebScreen(0,"Capture Home Screen")
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
wait (45)
Call CaptureWebScreen(0,"Capture Home Screen")

Call ClickWebButton(0,"","sapUshellAppTitle","Home","DIV","",False)

'Call ClickWebElement(0, "", "H1", "Home", "sapUshellHeadTitle", 0, False)
wait (15)
Call CaptureWebScreen(0,"Capture Production Screen")

Call ClickWebElement(0, "", "DIV", "Home Page Apps", "sapMSLITitleOnly", 0, False)
wait (15)
Call CaptureWebScreen(0,"Capture Production Screen")

Call ClickWebElement(0, "", "LI", "Production", "sapMLIB sapMLIB-CTX.*", 0, False)
wait (15)
Call CaptureWebScreen(0,"Capture Production Screen")

Call ClickWebElement(0, "", "LI", "Confirm Production Operations", "sapMLIB sapMLIB-CTX sapMLIBShowSeparator.*", 0, False)
wait (10)
Call CaptureWebScreen(0,"Capture Goods Receipt Screen")

Call WebEditSearch(0,"WebEdit",".*field0-I","INPUT","",DT_ORDER_NUMBER)
Call CaptureWebScreen(0,"Capture Goods Receipt Screen")


Call ClickSAPUIButton(0,"","BUTTON","Confirm","sapMBtnBase sapMBtn sapMBarChild","",False)
Call CaptureWebScreen(0,"Capture Goods Receipt Screen")

Call ClickWebElement(0, "", "SPAN", "Destination Tank:.*", "sapMObjectAttributeText", "", False)

Call SelectSAPUIRadioButton(0, "", "DIV", ".*finalConfirm_radioButton", "Final", "sapMRb.*", "", "On", False)

Call CaptureWebScreen(0,"Capture Goods Receipt Screen")
Call ClickSAPUIButton(0, "", "BUTTON", "Save", "sapMBtnBase.*", "", False)


'Call ClickWebElementById(0, "", "SPAN", ".*button29-inner", "WebElement", "", False)
Call CaptureWebScreen(0,"Capture Goods Receipt Screen")


CAll LogoffSRM(0)
Call FinalStatus()

