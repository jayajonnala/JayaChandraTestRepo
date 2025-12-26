
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_116_SRM Standard SC Limit Item on Cost Center (all opcos)_P1
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

gstrTestCaseName = "Test_116_SRM Standard SC Limit Item on Cost Center (all opcos)_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''

'''Close All Browser
Call CloseAllBrowsers()

''launch adn Login SRM Application
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
CAll SetWebEditFrameLogicalName(0, "Personal Object Work List", "Name of Shopping Cart", "text", 0, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0,"")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "SC number", "text", 0, "")
Call CaptureWebScreen(0,"Shopping Cart Monitor1")
'Search the Shopping cart name and Created at 
'Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Name of Shopping Cart", "text", 0, DT_GMLJUWLMAINVIEWFILTER_SUBJECT)
'Enter the Description
Wait(1)
'Call CaptureWebScreen(0,"Shopping Cart Description")

Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0, ConvertDAte(Date()-5))
'Enter the Description
Wait(1)
'Call CaptureWebScreen(0,"From Date value")

Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2,ConvertDAte(Date()+90))
'Enter the Description
Wait(1)
'Call CaptureWebScreen(0,"To Date value")

Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "SC number", "text", 0, DT_NUMBER)

'Enter the Description
Wait(1)
Call CaptureWebScreen(0,"To Date value")

Call ClickFrameSAPButton(0,"Personal Object Work List","Apply","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Apply")

Call ClickWebElementFrame(0, "Personal Object Work List", "SPAN", DT_NUMBER, 0)
Wait 10
Call CaptureWebScreen(1,"Display Shopping Cart")

Call ClickWebElementFrame(1, "Display Shopping Cart", "SPAN", "Purchase order created", 0)
Call CaptureWebScreen(1,"Purchase order created")

Call ClickWebElementFrame(1, "Shopping Cart", "SPAN", "Purchase Order", 0)
Call CaptureWebScreen(1,"View created Purchase order")

'''Call GetValueWebElementFrame(1, "Floor Plan Manager application for OIF","WD0365", "SPAN", "lsTextView.*", "DT_PONUM_OUTPUT")
'''
'''Call GetValueWebEdit(1, "Floor Plan Manager application for OIF", "lsField.*", "INPUT", "WD036A", 0, "DT_ITEM_NUMBER")

Call GetValueWebElementSAPFrame(1, "Floor Plan Manager application for OIF", "/SAPSRM/WDC_UI_DO_HISTORY\.ID_D5F1D826D57548788CFE59E2AB006028:V_BACKEND_PO\.PO_NO", "lsTextView.*", "DT_PONUM_OUTPUT")
Call GetValueWebEditLogicalNameSapFrame(1, "Floor Plan Manager application for OIF", "lsField.*", "INPUT", "Item Number", 0, "DT_ITEM_NUMBER")


Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","Close Window.*","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Show my Tasks")

Call VerifyFrameWebElement(1, "", "Shopping Cart", "SPAN", "ApprovedStatus", "lsTextView.*", 0, False)
Call CaptureWebScreen(1,"View created Purchase order")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickWebElementFrame(1, "Shopping Cart", "SPAN",DT_PO_ITEM, 0)
Wait 30
Browser("CreationTime:=2").FullScreen
Call CaptureWebScreen(2,"Capture screen:Purchase Order Overview")

Call VerifySAPEditFrame(2,"DG SRM  PO.*","Purchasing Document","text",0,DT_PONUM)
Wait 10
Call CaptureWebScreen(2,"Capture screen:Purchase Order Overview")
Browser("CreationTime:=2").Close
Wait 30

'''Call ClickFrameSAPButton(1,"Shopping Cart","Close","SPAN",0)
Call ClickFrameSAPButton(1,"Shopping Cart","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Home Screen")

'''------------------------'Log Off From Applicaton--------------------------------

Call LogOffSRM(0)
Call FinalStatus ()

