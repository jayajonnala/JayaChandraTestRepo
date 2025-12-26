
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


gstrTestCaseName = "TC_04_Test_Creation Goods receipts in FIORI_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
'''''Close All Browser
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

Call ClickWebElement(0, "", "LI", "Goods ReceiptAg. Inbound Delivery", "sapMLIB sapMLIB-CTX sapMLIBShowSeparator.*", 0, False)
wait (10)
Call CaptureWebScreen(0,"Capture Goods Receipt Screen")

Call SetWebEdit(0, "", "WebEdit", "search", 0, DT_DESCRIPTION, False)
Call CaptureWebScreen(0,"Capture Search Data Screen")
Call ClickWebElementById(0, "", "DIV", "__xmlview0--searchField-search", "sapMSFS sapMSFB", 0, False)
Wait 15
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call GetWebElementValueHtmlID(0, DT_PO_NUMBER, "sapMText sapUiSelectable.*", "SPAN", "", "DT_HTML_ID")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickWebButtonHtmlId(0, "", "sapUiIcon sapUiIconMirrorInRTL.*", ".*__clone"&DT_CLONEVALUE&".*", "SPAN", "", False)
Call VerifyWebElementById(0, "", "SPAN", DT_INBOUND_DEL, "__status.*", "", False)

Call SetWebEditByID(0, "WebEdit", ".*inputAadReferenceCode-inner", "INPUT", "", DT_AAD_REF_CODE)
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call ClickWebElementById(0, "", "SPAN", ".*buttonSelectTank-inner", "sapMBtnInner sapMBtnHoverable.*", 0, False)
Wait 15
Call CaptureWebScreen(0,"Capture Search results Data Screen")

'Call ClickWebElement(0, "", "LI", "0509 Tank.*", "sapMLIB sapMLIB-CTX.*", 0, False)
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call ClickWebElement(0, "", "SPAN", "Save", "sapMBtnInner sapMBtnHoverable.*", 0, False)
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call GetWebElementValueInnerText(0, "__text.*", "sapMText sapUiSelectable sapMTextMaxWidth textBold", "SPAN", 0,"", "DT_MESSAGE")
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call ClickWebElement(0, "", "SPAN", "Ok", "sapMBtnInner sapMBtnHoverable.*", 0, False)
wait 15
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call GetWebElementValueInnerText(0, "__text.*", "sapMText sapUiSelectable sapMTextMaxWidth sapMMsgBoxText", "SPAN", 0,"", "DT_FINAL_MESSAGE")
Call CaptureWebScreen(0,"Capture Search results Data Screen")

Call ClickWebElement(0, "", "SPAN", "OK", "sapMBtnInner sapMBtnHoverable sapMFocusable sapMBtnText sapMBtnDefault", 0, False)
wait 15
Call CaptureWebScreen(0,"Capture Search results Data Screen")

CAll LogoffSRM(0)
Call FinalStatus()

