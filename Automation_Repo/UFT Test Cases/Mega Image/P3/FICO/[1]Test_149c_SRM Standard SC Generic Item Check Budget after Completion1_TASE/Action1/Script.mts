
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_149c_SRM Standard SC Generic Item Check Budget after Completion1
'.................Author : TCS 	   :
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

gstrTestCaseName = "Test_149c_SRM Standard SC Generic Item Check Budget after Completion1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'''''Close All Browser
Call CloseAllBrowsers()

''launch and Login SRM Application
'Call LaunchSAPWebApplication(DT_SAPURL)
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)
'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")
wait (5)

''Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","My InboxExpert Mode","sapMGTHdrContent.*",0,False)
wait (5)
Call CaptureWebScreen(0,"Capture Expert mode screen")
Call WebEditSearch(0, "WebEdit", ".*field0-I", "INPUT", 0, DT_GMLJUWLMAINVIEWFILTER_SUBJECT)

Call CaptureWebScreen(0,"Shopping Cart Description")

Call ClickWebElement(0, "", "DIV", "", "sapMSFS sapMSFB", 0, False)

Call ClickLink(0, "", "sapMLnk sapMLnkMaxWidth", DT_SC_NAME, False)
Wait 10
Call CaptureWebScreen(1,"Shopping subject Description")

Call ClickFrameSAPButton(1,"Approve Shopping Cart","Edit","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Edit button")
Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","OK","DIV",1)
Call ClickFrameSAPButton(1,"Change Shopping Cart and Proceed","Details","DIV",0)
Wait 15
Call CaptureWebScreen(1,"Capture Details Screen")

Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Next Item","IMG",0)
Wait 15
Call CaptureWebScreen(1,"Capture Next Item Screen")


Call ClickWebElementFrame(1, "Shopping Cart and Proceed", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Details","DIV",1)
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing

''Call ClickWebElementFrameNoInnertext(1, "Shopping Cart and Proceed", "SPAN", "WD06B0-r", 0)
''Wait 15
''Call ClickWebElementFrame(1, "Shopping Cart and Proceed", "TR", "Asset", 0)


''Call SelectValSAPList(1, "Shopping Cart and Proceed","Account Assignment Category", "WD0932","Asset")
'' SelectValSAPList(creationTime, listName, listHtmlTag, listIndex, listValue)
''Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"Asset")
''Call SelectValFrameSAPList(1, "Shopping Cart and Proceed","Account Assignment Category", "","INPUT",0,"Asset")

Call ClickSAPList(1, "Account Assignment Category", "INPUT",1)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 15

Call SetWebEditFrameByTitle(1, "Shopping Cart and Proceed", "Cost Center", "text", 0, DT_COSTCENTER)
Call SetWebEditFrameByTitle(1, "Shopping Cart and Proceed", "Enter the WBS Element", "text", 0, DT_WBS_ELEMENT_)

WAit 15
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Create Asset Master","DIV",0)
Wait 15
'Call CaptureWebScreen(0,"Capture Next Item Screen")

Call ClickWebElementFrame(1, "Shopping Cart and Proceed", "DIV", "Item Data", 0)
Call SetWebEditFrameLogicalName(1,"Shopping Cart and Proceed","Plant / Location","text",1,7625)
Call CaptureWebScreen(1,"Capture Next Item Screen")

Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Approve","DIV",0)
Wait 15
Call CaptureWebScreen(1,"Approve button")
'Call VerifyWebElement(1, "", "SPAN", DT_, "lsTextView lsTextView.*", 0, FAlse)
Wait 5
Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Refresh")
Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 10

'''------------------------'Log Off From Applicaton--------------------------------

Call LogOffSRM(0)
Call FinalStatus ()

''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

