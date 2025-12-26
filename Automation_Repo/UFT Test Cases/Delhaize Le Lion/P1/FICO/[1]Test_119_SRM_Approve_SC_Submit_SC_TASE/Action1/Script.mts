
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_119_SRM_Approve_SC_Submit_SC_TASE
'.................Author : TCS 	   
'................ Creation Date    
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

gstrTestCaseName = "Test_119_SRM_Approve_SC_Submit_SC_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
'Close All Browser
Call CloseAllBrowsers()

''launch adn Login SRM Application
'Call LaunchSAPWebApplication(DT_SAPURL)
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")
''Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","My InboxExpert Mode","sapMGTHdrContent.*",0,False)
wait (5)
Call CaptureWebScreen(0,"Capture Expert mode screen")
Call WebEditSearch(0, "WebEdit", ".*field0-I", "INPUT", 0, DT_SC_NAME)

Call CaptureWebScreen(0,"Shopping Cart Description")

Call ClickWebElement(0, "", "DIV", "", "sapMSFS sapMSFB", 0, False)
Call CaptureWebScreen(0,"Capture screen:Refresh")
CAll ClickLink(0, "", "sapMLnk sapMLnkMaxWidth", DT_GMLJUWLMAINVIEWFILTER_SUBJECT, False)
Wait 10
Call CaptureWebScreen(1,"Capture Subject")


Call CaptureWebScreen(1,"Shopping subject Description")
Call ClickFrameSAPButton(1,"Approve Shopping Cart","Submit","DIV",0)
Wait 10
'Call VerifyWebElement(1, "", "SPAN", DT_, "lsTextView lsTextView.*", 0, FAlse)
Wait 5
Call CaptureWebScreen(1,"Submit button")

Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Refresh")
CAll VerifyFrameWebElement(1, "", "Document:", "SPAN", DT_MESSAGE, "lsTextView.*", 0, False)

Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Capture screen:Refresh")
'''------------------------'Log Off From Applicaton--------------------------------

Call LogOffSRM(0)
Call FinalStatus ()

