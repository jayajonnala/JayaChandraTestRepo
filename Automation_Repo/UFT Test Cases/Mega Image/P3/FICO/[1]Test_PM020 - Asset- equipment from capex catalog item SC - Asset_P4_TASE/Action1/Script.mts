
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_PM020 - Asset- equipment from capex catalog item SC - Asset_P4
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

gstrTestCaseName = "Test_PM020 - Asset- equipment from capex catalog item SC - Asset_P4_TASE"
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
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)
'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
wait (5)
Call CaptureWebScreen(0,"Capture Home Screen")


''Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","My InboxExpert Mode","sapMGTHdrContent.*",0,False)
wait (15)
Call CaptureWebScreen(0,"Capture Expert mode screen")
Call WebEditSearch(0, "WebEdit", ".*field0-I", "INPUT", 0, DT_SCNAME)
Call CaptureWebScreen(0,"Shopping Cart Description")
Call ClickWebElement(0, "", "DIV", "", "sapMSFS sapMSFB", 0, False)
Call ClickLink(0, "", "sapMLnk sapMLnkMaxWidth", DT_GMLJUWLMAINVIEWFILTER_SUBJECT, False)
Wait 10
Call CaptureWebScreen(1,"Shopping subject Description")

Call ClickFrameSAPButton(1,"Approve Shopping Cart","Edit","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Edit button")

Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","OK","DIV",1)

Call ClickFrameSAPButton(1,"Change Shopping Cart and Proceed","Details","DIV",0)
Wait 15
Call CaptureWebScreen(1,"Edit button")
Call ClickWebElementFrame(1, "Shopping Cart and Proceed", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Details","DIV",1)
Wait 15
Call CaptureWebScreen(1,"Details button")

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing

''''Call ClickWebElementFrameNoInnertext(1, "Shopping Cart and Proceed", "SPAN", "WD0698-btn", 0)
''''''Wait 15
''''CAll ClickWebElementFrame(1, "Shopping Cart and Proceed", "TR", "Asset", 0)

'Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"Asset")
Call ClickSAPList(1, "Account Assignment Category", "INPUT",1)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 15
Call CaptureWebScreen(1,"Edit button")
''''Call SetWebEditFrame(1, "Shopping Cart and Proceed", "WD0838", "text", 0, DT_WBS_ELEMENT_)
Call SetWebEditFrameByTitle(1, "Shopping Cart and Proceed", "Enter the WBS Element", "text", 0, DT_WBS_ELEMENT_)

WAit 5
Call CaptureWebScreen(1,"Edit button")
Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Create Asset Master","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Edit button")
Call GetValueFrameWebEdit(1, "Shopping Cart and Proceed", "lsField__input", "INPUT", "/SAPSRM/WDC_UI_DO_ACC\.ID_A56B0481584E314F46F1C5E908E8819E:V_DO_ACCOUNT_DETAIL\.ASSET", 0, "DT_ASSET_OUTPUT")

Call ClickFrameSAPButton(1,"Shopping Cart and Proceed","Approve","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Approve button")
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

