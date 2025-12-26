
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_125_SRM Technical SC Catalog Items on Asset 2 Approvals (DLL-MI)_P1_TASE
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

gstrTestCaseName = "Test_125_SRM Technical SC Catalog Items on Asset 2 Approvals (DLL-MI)_P1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''Close All Browser
Call CloseAllBrowsers()

''launch adn Login SRM Application
''Call LaunchSAPWebApplication(DT_SAPURL)
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")


'''Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","My InboxExpert Mode","sapMGTHdrContent.*",0,False)
wait (5)
Call CaptureWebScreen(0,"Capture Expert mode screen")

Call WebEditSearch(0, "WebEdit", ".*field0-I", "INPUT", 0, DT_SUBJECT)
Call ClickWebElement(0, "", "DIV", "", "sapMSFS sapMSFB", 0, False)
Call CaptureWebScreen(0,"Capture screen:Refresh")

CAll ClickLink(0, "", "sapMLnk sapMLnkMaxWidth", DT_GMLJUWLMAINVIEWFILTER_SUBJECT, False)
Wait 10
Call CaptureWebScreen(1,"Capture Subject")


Call ClickFrameSAPButton(1, "Approve Shopping Cart", "Edit", "DIV", 0)
Wait 5
Call CaptureWebScreen(1,"Approve Shopping Cart")

Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","OK","DIV",1)
Wait 5
Call CaptureWebScreen(1,"Capture Ok Button")


Call ClickFrameSAPButton(1, "Change Shopping Cart and Proceed", "Details", "DIV", 0)
Wait 30
Call CaptureWebScreen(1,"Capture 1stdetails button")
Call ClickWebElementFrame(1, "Shopping Cart and Proceed", "DIV", "Account Assignment", 0)

Call ClickFrameSAPButton(1, "Shopping Cart and Proceed", "Details", "DIV", 1)
Wait 5

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{END}",1
Set wsh= nothing
Call CaptureWebScreen(1,"CaptureAccount assignment")

Wait 10
'Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"Asset")
'Wait 10
'Call CaptureWebScreen(1,"Capture Asset1")
Call ClickSAPList(1, "Account Assignment Category", "INPUT",1)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5


Call SetSAPEdit(1,"WBS Element", "text", 0, DT_WBS_ELEMENT_01)
Wait 5
Call CaptureWebScreen(1,"Capturefilled details")

Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Create Asset Master","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Create Asset Master")
Wait 2
Call GetValueSAPEdit(1,"Asset","/SAPSRM/WDC_UI_DO_ACC\.ID_A56B0481584E314F46F1C5E908E8819E:V_DO_ACCOUNT_DETAIL\.ASSET", "DT_ASSET_01_OUTPUT")
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Next Item","IMG",0)
Wait 10

'Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"Asset")
'Wait 10
Call ClickSAPList(1, "Account Assignment Category", "INPUT",1)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5
Call CaptureWebScreen(1,"Capture Asset2")

Call SetSAPEdit(1,"WBS Element", "text", 0, DT_WBS_ELEMENT_02)

Call CaptureWebScreen(1,"Capturefilled details")
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Create Asset Master","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Create Asset Master2")

Call GetValueSAPEdit(1,"Asset","/SAPSRM/WDC_UI_DO_ACC\.ID_A56B0481584E314F46F1C5E908E8819E:V_DO_ACCOUNT_DETAIL\.ASSET", "DT_ASSET_02_OUTPUT")

Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Check","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture Check screen")
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Approve","DIV",0)
Wait 10
Call CaptureWebScreen(1,"CaptureApprove screen")

Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture Refresh screen")
CAll VerifyFrameWebElement(1, "", "Document:", "SPAN", DT_AWAITING_APPROVAL, "lsTextView.*", 0, False)

Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Capture close button")

'''------------------------'Log Off From Applicaton--------------------------------

Call LogOffSRM(0)
Call FinalStatus ()





