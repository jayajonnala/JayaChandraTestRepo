
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_149c_SRM Standard SC Generic Item Check Budget after Completion2
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

gstrTestCaseName = "Test_149c_SRM Standard SC Generic Item Check Budget after Completion2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''Close All Browser
Call CloseAllBrowsers()

''launch and Login SRM Application
'Call LaunchSAPWebApplication(DT_SAPURL)
Call LaunchSAPWebApplicationEdge(DT_SAPURL)

'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")

'Click on Shopping Cart Monitor
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call ClickWebButton(0,"",".*","Shopping Cart Monitor","DIV",0,False)
wait (5)
Call CaptureWebScreen(0,"Capture Home Screen")

'Call SetWebEditFrame(0, "Personal Object Work List", "WD57", "text", 0, "")
'Call SetWebEditFrame(0, "Personal Object Work List", "WD0151", "text", 0,"")
'Call SetWebEditFrame(0, "Personal Object Work List", "WD0155", "text", 0, "")
'Call SetWebEditFrame(0, "Personal Object Work List", "WD68", "text", 0, "")

CAll SetWebEditFrameLogicalName(0, "Personal Object Work List", "Name of Shopping Cart", "text", 0, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0,"")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "SC number", "text", 0, "")

'Search the Shopping cart name and Created at 
CAll SetWebEditFrameLogicalName(0, "Personal Object Work List", "Name of Shopping Cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
'Call SetWebEditFrame(0, "Personal Object Work List", "WD57", "text", 0, DT_NAME_OF_SHOPPING_CART)
'Enter the Description
Wait(1)
'Call SetWebEditFrame(0, "Personal Object Work List", "WD0151", "text", 0, ConvertDAte(Date()-5))
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0,ConvertDAte(Date()-5))
'Enter the Description
Wait(1)
'Call SetWebEditFrame(0, "Personal Object Work List", "WD0155", "text", 0, ConvertDAte(Date()+90))
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2, ConvertDAte(Date()+90))
'Enter the Description
Wait(1)
'Call SetWebEditFrame(0, "Personal Object Work List", "WD68", "text", 0, DT_SC_NUMBER)
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "SC number", "text", 0, DT_SC_NUMBER)
'Enter the Description
Wait(1)
Call CaptureWebScreen(0,"To Date value")

Call ClickFrameSAPButton(0,"Personal Object Work List","Apply","DIV",0)
Wait 30
Call CaptureWebScreen(0,"Apply Button")

Call ClickWebElementFrame(0, "Personal Object Work List", "SPAN", DT_SC_NUMBER, 0)
Wait 20
Call CaptureWebScreen(0,"To Date value")

Call ClickWebElementFrame(1, "Display Shopping Cart", "SPAN", "Purchase order created", 0)
Wait 10
Call ClickWebElementFrame(1, "Shopping Cart", "SPAN", "Purchase Order", 0)
WAit 20
Call CaptureWebScreen(1,"View created Purchase order")

Call GetValueWebElementSAPFrame(1, "Floor Plan Manager application for OIF", "/SAPSRM/WDC_UI_DO_HISTORY\.ID_D5F1D826D57548788CFE59E2AB006028:V_BACKEND_PO\.PO_NO", "lsTextView.*", "DT___AGIM0U112021_OUTPUT")
Call GetValueWebEditSapFrame(1, "Floor Plan Manager application for OIF", "lsField.*", "INPUT", "/SAPSRM/WDC_UI_DO_HISTORY\.ID_D5F1D826D57548788CFE59E2AB006028:V_BACKEND_PO\.ITEM_NO", 0, "DT_ITEM_NUMBER")

'Call ClickFrameSAPButton(1,"Floor Plan Manager application for OIF","Close Window","SPAN",0)
Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","Close Window.*","DIV",0)
Wait 15
Call CaptureWebScreen(1,"Capture screen:Show my Tasks")

Call VerifyFrameWebElement(1, "", "Shopping Cart", "SPAN", DT_STATUS, "lsTextView.*", 0, False)


Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickWebElementFrame(1, "Shopping Cart", "SPAN",DT_PO_ITEM, 0)
Wait 30
Browser("CreationTime:=2").FullScreen
Call CaptureWebScreen(2,"Capture screen:Purchase Order Overview")

Call VerifySAPEditFrame(2,"DG SRM  PO.*","Purchasing Document","text",0,DT_PO)
Wait 10
Browser("CreationTime:=2").Close
Wait 30

Call CaptureWebScreen(1,"Capture screen:Show my Tasks")
'Call ClickFrameSAPButton(1,"Shopping Cart","Close","SPAN",0)
Call ClickFrameSAPButton(1,"Shopping Cart","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Home Screen")
''
''''------------------------'Log Off From Applicaton--------------------------------
Call LogoffSRM(0)
Call FinalStatus()

