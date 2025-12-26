
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_125_SRM Technical SC Catalog Items on Asset 2 Approvals (DLL-MI)_TASE
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

gstrTestCaseName = "Test_125_SRM Technical SC Catalog Items on Asset 2 Approvals (DLL-MI)_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
''''Close All Browser
Call CloseAllBrowsers()

''launch adn Login SRM Application
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
wait (5)
Call CaptureWebScreen(0,"Capture Home Screen")


'Click on Create Professional Shopping Cart
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call CaptureWebScreen(0,"Shopping & Receiving")
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
wait (5)
Browser("CreationTime:=1").Fullscreen
Call CaptureWebScreen(1,"Create Professional Shopping Cart")
'Call SetWebEditFrame(1, "Create Shopping Cart", "WDA6", "text", 0, DT_NAME_OF_SHOPPING_CART)
Call SetSAPEdit(1, "Name of shopping cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
Wait(1)
Call CaptureWebScreen(1,"Shopping Cart Description")

Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)

Call CaptureWebScreen(1,"Click Add Item")
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)
Call CaptureWebScreen(1,"Click Add Item1")
Call ClickWebElementSAPFrame(1, "Floor Plan Manager application for OIF", ".*text", "DLL SRM-MDM PreProduction BE - All", "lsLink.*", 0)

Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMCJH\.BasicView\.KeywordCriteriaInputField", "text", 0, DT_KEYWORD)
Call CaptureWebScreen(1,"Enter Search KeyWord")

Call ClickFrameSAPButton(1,"ProcurementCatalog7Ehp1","Search","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Search Result")
'Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorQuantity\.0", "text", 0, 1)
Call SetCellDataSAPTableFrame(1,"ProcurementCatalog7Ehp1","urST3BdBrd urST3Bd urFontStd","TABLE","Add to CartAdd to Shopping ListCompareResults ViewResults Per PageImage Off",0,2,2,1)
Call ClickWebElementFrameNoInnertext(1, "ProcurementCatalog7Ehp1", "SPAN", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorCheckBox\.0", 0)
Call CaptureWebScreen(1,"Item Selection1")

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Wait 10
Call CaptureWebScreen(1,"Item Selection display1")
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 30
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Call CaptureWebScreen(1,"Account Assignment1")
''''Call SetWebEditFrame(1, "Shopping Cart", "WD0AE1", "text", 0, DT_COSTCENTER)

Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",2,DT_COSTCENTER)
Wait(1)
Call CaptureWebScreen(1,"Shopping Cart Description")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Item Data", 0)
Call CaptureWebScreen(1,"Shopping Cart Description")
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Plant / Location","text",1,6634)


Call ClickFrameSAPButton(1,"Shopping Cart","Add Item","DIV",0)
Call CaptureWebScreen(1,"Click Add Item")
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)
Call CaptureWebScreen(1,"Click Add Item2")
Call ClickWebElementSAPFrame(1, "Floor Plan Manager application for OIF", ".*text", "DLL SRM-MDM PreProduction BE - All", "lsLink.*", 0)

Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMCJH\.BasicView\.KeywordCriteriaInputField", "text", 0, DT_KEYWORD_OCC1)
Call CaptureWebScreen(1,"Enter Search KeyWord")

Call ClickFrameSAPButton(1,"ProcurementCatalog7Ehp1","Search","DIV",0)
Call CaptureWebScreen(1,"Search Result2")
'Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorQuantity\.0", "text", 0, 1)
Call SetCellDataSAPTableFrame(1,"ProcurementCatalog7Ehp1","urST3BdBrd urST3Bd urFontStd","TABLE","Add to CartAdd to Shopping ListCompareResults ViewResults Per PageImage Off",0,2,2,1)
Call ClickWebElementFrameNoInnertext(1, "ProcurementCatalog7Ehp1", "SPAN", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorCheckBox\.0", 0)
Call CaptureWebScreen(1,"Item Selection2")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Wait 10
Call CaptureWebScreen(1,"Item Selection display2")
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 30
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)

Call ClickFrameSAPButton(1,"Shopping Cart","Next Item","IMG",0)
WAit 5
Call CaptureWebScreen(1,"Account Assignment2")
''''Call SetWebEditFrame(1, "Shopping Cart", "WD10C8", "text", 0, DT_COSTCENTER_OCC1)
'''' SetSAPEdit(creationTime, SAPAttachedtext, SAPEditTexttype, SAPEditIndex, val)

Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",2,DT_COSTCENTER_OCC1)
Wait(5)
Call CaptureWebScreen(1,"Shopping Cart Description")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Item Data", 0)
Call CaptureWebScreen(1,"Shopping Cart Description")
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Plant / Location","text",1,6634)

Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Capture screen:Check Out")


Call ClickFrameSAPButton(1,"Shopping Cart","Order","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Capture screen:Complete order")

'''Call GetValueWebElementFrame(1, "Display Document:", "WD11FC", "SPAN", "ls.*", "DT__OCC2_OUTPUT")
'Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_LIST_TEXTVIEW\.1", "SPAN", "ls.*", "DT__OCC2_OUTPUT")
'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", DT__OCC2_OUTPUT, "ls.*", 0, False)

Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_1", "DIV", "ls.*", "DT__OCC2_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT__OCC2_OUTPUT, "ls.*", 0, False)

Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Capture screen:Refresh")

Call ClickWebElementFrame(1, "Document:", "SPAN", "Display / Edit Agents", 0)
Wait 60
Call CaptureWebScreen(1,"Capture screen:Approval Flow")

'Call GetValueWebElementSAPFrame(1, "Floor Plan Manager application for OIF", "WD0C8C-text", "ls.*", "DT_PROCESSOR")
Call ClickSAPFrameSAPButton(1, "Floor Plan Manager application for OIF", "OK", "DIV", 0)
Wait 20
Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Capture screen:Close button")
Browser("CreationTime:=0").Maximize
Call LogoffSRM(0)
Call FinalStatus()


'
''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

