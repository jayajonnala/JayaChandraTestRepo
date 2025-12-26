
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_119_SRM_Complete_Approve_SC
'.................Author : TCS 	  
'................ Creation Date    : 
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

gstrTestCaseName = "Test_119_SRM_Complete_Approve_SC_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'''''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''Close All Browser
Call CloseAllBrowsers()

'''''launch adn Login SRM Application
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")


'''''Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","My InboxExpert Mode","sapMGTHdrContent.*",0,False)
wait (5)
Call CaptureWebScreen(0,"InboxExpert Mode")
Call WebEditSearch(0, "WebEdit", ".*field0-I", "INPUT", 0, DT_SC_NAME)
Call ClickWebElement(0, "", "DIV", "", "sapMSFS sapMSFB", 0, False)
Call CaptureWebScreen(0,"SearchResult")
CAll ClickLink(0, "", "sapMLnk sapMLnkMaxWidth", DT_GMLJUWLMAINVIEWFILTER_SUBJECT, False)
Wait 15
Call CaptureWebScreen(1,"Capture Subject")


Call ClickFrameSAPButton(1, "Approve Shopping Cart", "Edit", "DIV", 0)
Wait 5
Call CaptureWebScreen(1,"Capture Edit")
''Call ClickSAPButtonSAPFrame(1, "Floor Plan Manager application for OIF", "lsButton__text urBtnCntTxt", "OK", "DIV", 0)
Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","OK","DIV",1)
Wait 5
Call CaptureWebScreen(1,"Capture Ok Button")
Call ClickFrameSAPButton(1, "Change Shopping Cart and Proceed", "Details", "DIV", 0)
Wait 5
Call CaptureWebScreen(1,"Capture 1stdetails button")
Call ClickWebElementFrame(1, "Shopping Cart and Proceed", "DIV", "Account Assignment", 0)
Call CaptureWebScreen(1,"CaptureAccount assignemnet")
Call ClickFrameSAPButton(1, "Shopping Cart and Proceed", "Details", "DIV", 1)
Wait 5

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{END}",1
Set wsh= nothing
Call CaptureWebScreen(1,"Account Assignment1")

Wait 10
''''Call ClickWebElementFrameHtmlID(1, "Shopping Cart and Proceed", "SPAN", "WD068C-btn",0)
''''Wait 2
''''Call ClickWebElementFrameHtmlID(1, "Shopping Cart and Proceed", "SPAN", "WD068C-btn",0)
''''Wait 2
''''Call ClickWebElementFrameHtmlID(1, "Shopping Cart and Proceed", "TR", "WD0692", 0)
'Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"OPEX WBS element")
'Wait 10
Call ClickSAPList(1, "Account Assignment Category", "INPUT",0)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{END}",1
Wait(2)
wsh.SendKeys "{UP}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5

Call CaptureWebScreen(1,"Capture Opex WBS Element")

''''No need of updating the cost center value, it's already update in the 1st test case.Hence commenting the below line of code
''Call SetWebEditFrame(1, "Shopping Cart and Proceed", "WD069D", "text", 0, DT_COSTCENTER)

''''Call SetWebEditFrame(1, "Shopping Cart and Proceed", "WD0723", "text", 0, DT_WBS_ELEMENT_)
Call SetSAPEdit(1,"WBS Element", "text", 0, DT_WBS_ELEMENT_)
Wait 5
Call CaptureWebScreen(1,"Capturefilled details")

Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Next Item","IMG",0)
Wait 10
Call CaptureWebScreen(1,"Capture second item")

''''Call ClickWebElementFrameHtmlID(1, "Shopping Cart and Proceed", "SPAN", "WD068C-btn", 3)
''''Wait 2
''''Call ClickWebElementFrameHtmlID(1, "Shopping Cart and Proceed", "TR", "WD0692", 0)
'Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"OPEX WBS element")
'Wait 15
Call ClickSAPList(1, "Account Assignment Category", "INPUT",0)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{END}",1
Wait(2)
wsh.SendKeys "{UP}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5
Call CaptureWebScreen(1,"Capture second item")

''''No need of updating the cost center value, it's already update in the 1st test case.Hence commenting the below line of code
'''''Call SetWebEditFrame(1, "Shopping Cart and Proceed", "WD06A6", "text", 0, DT_COSTCENTER_OCC1)

'''Call SetWebEditFrame(1, "Shopping Cart and Proceed", "WD0798", "text", 0, DT_WBS_ELEMENT__OCC1)
Call SetSAPEdit(1,"WBS Element", "text", 0, DT_WBS_ELEMENT__OCC1)
Wait 5
Call CaptureWebScreen(1,"Capture filled details")

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

