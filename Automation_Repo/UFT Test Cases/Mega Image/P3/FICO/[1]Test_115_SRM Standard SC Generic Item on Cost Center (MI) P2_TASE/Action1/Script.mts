
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_115_SRM Standard SC Generic Item on Cost Center (MI) P2
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



gstrTestCaseName = "Test_115_SRM Standard SC Generic Item on Cost Center (MI) P2 _TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
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
Call ClickWebButton(0,"",".*","Shopping Cart Monitor","DIV",0,False)
wait (5)
Call CaptureWebScreen(0,"Capture Home Screen")

 'Clearing the previus search results.
''Call SetWebEditFrame(0, "Personal Object Work List", "WD57", "text", 0, "")
''Call SetWebEditFrame(0, "Personal Object Work List", "WD0151", "text", 0, "")
''Call SetWebEditFrame(0, "Personal Object Work List", "WD0155", "text", 0, "")
''Call SetWebEditFrame(0, "Personal Object Work List", "WD68", "text", 0, "")
Call SetWebEditFrame(0, "Personal Object Work List", "WDR_SELECT_OPTIONS\.ID_6475B2BC42BC8CF92BE5404B1931A2F8:SELECTION_SCREEN\.ILP_DESCRH", "text", 0, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0, "")
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2, "")
Call SetWebEditFrame(0, "Personal Object Work List", "WDR_SELECT_OPTIONS\.ID_6475B2BC42BC8CF92BE5404B1931A2F8:SELECTION_SCREEN\.ILS_OBJID", "text", 0, "")

'Search the Shopping cart name and Created at
''Call SetWebEditFrame(0, "Personal Object Work List", "WD57", "text", 0, DT_NAME_OF_SHOPPING_CART)
Call SetWebEditFrame(0, "Personal Object Work List", "WDR_SELECT_OPTIONS\.ID_6475B2BC42BC8CF92BE5404B1931A2F8:SELECTION_SCREEN\.ILP_DESCRH", "text", 0, DT_NAME_OF_SHOPPING_CART)
'Enter the Description
Wait(1)

''Call SetWebEditFrame(0, "Personal Object Work List", "WD0151", "text", 0, ConvertDAte(Date()-5))
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "Created on", "text", 0, ConvertDAte(Date()-5))
Wait(1)

'Call SetWebEditFrame(0, "Personal Object Work List", "WD0155", "text", 0, ConvertDAte(Date()+90))
Call SetWebEditFrameLogicalName(0, "Personal Object Work List", "To", "text", 2, ConvertDAte(Date()+90))
Wait(1)

''Call SetWebEditFrame(0, "Personal Object Work List", "WD68", "text", 0, DT_SC_NUMBER)
Call SetWebEditFrame(0, "Personal Object Work List", "WDR_SELECT_OPTIONS\.ID_6475B2BC42BC8CF92BE5404B1931A2F8:SELECTION_SCREEN\.ILS_OBJID", "text", 0, DT_SC_NUMBER)
Wait(1)

Call CaptureWebScreen(0,"DT_SC_NUMBER")

Call ClickFrameSAPButton(0,"Personal Object Work List","Apply","DIV",0)
Wait 20
Call CaptureWebScreen(0,"Clicking on Apply BUtton")

Call ClickWebElementFrame(0, "Personal Object Work List", "SPAN", DT_SC_NUMBER, 0)
Wait 30

Call ClickWebElementFrame(1, "Display Shopping Cart", "SPAN", "Purchase order created", 0)
Wait 10

Call ClickWebElementFrame(1, "Shopping Cart", "SPAN", "Purchase Order", 0)

Call CaptureWebScreen(1,"View created Purchase order")

'Call GetValueWebElementSAPFrame(1, "Floor Plan Manager application for OIF", "WD0337", "lsTextView.*", "DT_PONUM_OUTPUT")
Call GetValueWebElementSAPFrame(1, "Floor Plan Manager application for OIF", "/SAPSRM/WDC_UI_DO_HISTORY\.ID_D5F1D826D57548788CFE59E2AB006028:V_BACKEND_PO\.PO_NO", "lsTextView.*", "DT_PONUM_OUTPUT")

'Call GetValueWebEditSapFrame(1, "Floor Plan Manager application for OIF", "lsField.*", "INPUT", "WD033C", 0, "DT_ITEM_NUMBER")
Call GetValueWebEditSapFrame(1, "Floor Plan Manager application for OIF", "lsField.*", "INPUT", "/SAPSRM/WDC_UI_DO_HISTORY\.ID_D5F1D826D57548788CFE59E2AB006028:V_BACKEND_PO\.ITEM_NO", 0, "DT_ITEM_NUMBER")
'
''Call ClickFrameSAPButton(1,"Floor Plan Manager application for OIF","Close Window","SPAN",0)
Call ClickSAPFrameSAPButton(1,"Floor Plan Manager application for OIF","Close Window.*","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Show my Tasks")

'''Call ClickWebElementFrameHtmlID(1, "Shopping Cart", "SPAN", "WD02F0-text", 0)
'''need to add more lines of script
Call VerifyFrameWebElement(1, "", "Shopping Cart", "SPAN", "ApprovedStatus", "lsTextView.*", 0, False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickWebElementFrame(1, "Shopping Cart", "SPAN",Trim(Cstr(DT_PO_Item)), 0)
Wait 30
Browser("CreationTime:=2").FullScreen
Call CaptureWebScreen(2,"Capture screen:Purchase Order Overview")

Call VerifySAPEditFrame(2,"DG SRM  PO.*","Purchasing Document","text",0,Trim(Cstr(DT_PO_NUMBER)))
Wait 10
Browser("CreationTime:=2").Close

Call CaptureWebScreen(1,"Capture screen:Show my Tasks")
'Call ClickFrameSAPButton(1,"Shopping Cart","Close","SPAN",0)
Call ClickFrameSAPButton(1,"Shopping Cart","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Capture screen:Show my Tasks")


'
''''------------------------'Log Off From Applicaton--------------------------------
Call LogoffSRM(0)
Call FinalStatus()

