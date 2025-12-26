
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_114_SRM Standard SC Free Text on Cost Cent P2_Approve
'.................Author : TCS 	   :Kumar K
'................ Creation Date    : 12th Dec
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

gstrTestCaseName = "Test_114_SRM Standard SC Free Text on Cost Cent P2_Approve_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
'''Close All Browser
Call CloseAllBrowsers()

''launch and Login SRM Application
'Call LaunchSAPWebApplication(DT_SAPURL)
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")


'Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call CaptureWebScreen(0,"Capture Home Screen1")
Call ClickWebButton(0,"",".*","Shopping Cart Monitor","DIV",0,False)
wait (5)
Call CaptureWebScreen(0,"Shopping Cart Monitor")

' SetWebEditFrameLogicalName(creationTime, frameTitle, webeditLogicalName, webeditTexttype, webeditIndex, val)
CAll SetWebEditFrameLogicalName(0, "Personal Object Work List", "Name of Shopping Cart", "text", 0, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0,"")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "SC number", "text", 0, "")
Call CaptureWebScreen(0,"Shopping Cart Monitor1")


Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Name of Shopping Cart", "text", 0, DT_GMLJUWLMAINVIEWFILTER_SUBJECT)
Wait(1)
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0, ConvertDAte(Date()-5))
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2,ConvertDAte(Date()+90))
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "SC number", "text", 0, DT_SC_NUMBER)
Wait(1)
Call CaptureWebScreen(0,"To Date value")

Call ClickFrameSAPButton(0,"Personal Object Work List","Apply","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Apply Button")

Call ClickWebElementFrame(0, "Personal Object Work List", "SPAN", DT_SC_NUMBER, 0)
Wait 10
Call CaptureWebScreen(0,"To Date value")
Call ClickFrameSAPButton(1,"Display Shopping Cart","Show my Tasks","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Show my Tasks")

Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","Process Now","DIV","")
Call CaptureWebScreen(1,"Capture screen:Process Now")

Call ClickFrameSAPButton(1,"Approve Shopping Cart","Approve","DIV",0)
Wait 10
'Call ClickFrameSAPButton(1,"Display Document:","Refresh","SPAN",0)
Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 10
'Call ClickFrameSAPButton(1,"Document:","Refresh","SPAN",0)
'Wait 10
Call CaptureWebScreen(1,"Capture screen:Refresh")
Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 15
Call CaptureWebScreen(0,"Capture screen:Close button")

'''------------------------'Log Off From Applicaton--------------------------------

Call LogOffSRM(0)
Call FinalStatus ()

