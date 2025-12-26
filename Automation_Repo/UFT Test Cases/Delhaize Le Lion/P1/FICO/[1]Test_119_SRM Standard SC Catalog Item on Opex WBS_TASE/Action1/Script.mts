
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_119_SRM Standard SC Catalog Item on Opex WBS_TASE
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

gstrTestCaseName = "Test_119_SRM Standard SC Catalog Item on Opex WBS_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''Close All Browser
Call CloseAllBrowsers()

''launch adn Login SRM Application
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

'Call LoginSAPEdit(DT_SAPUSER,DT_SAPPASSWORD)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")

wait (5)
'Click on Create Professional Shopping Cart
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call CaptureWebScreen(0,"Shopping & Receiving")
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Wait 10
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
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)
Call CaptureWebScreen(1,"Click Add Item1")
Call ClickWebElementSAPFrame(1, "Floor Plan Manager application for OIF", ".*text", "DLL SRM-MDM Pre\. EU - BE - ALL", "lsLink.*", 0)

Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMCJH\.BasicView\.KeywordCriteriaInputField", "text", 0, DT_KEYWORD)
'Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMCJH\.BasicView\.KeywordCriteriaInputField", "", 0, DT_KEYWORD)
Call CaptureWebScreen(1,"Enter Search KeyWord")

Call ClickFrameSAPButton(1,"ProcurementCatalog7Ehp1","Search","DIV",0)
Wait 5
Call CaptureWebScreen(1,"SearchResult")
'Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorQuantity\.0", "text", 0, DT_QUANTITY)
Call SetCellDataSAPTableFrame(1,"ProcurementCatalog7Ehp1","urST3BdBrd urST3Bd urFontStd","TABLE","Add to CartAdd to Shopping ListCompareResults ViewResults Per PageImage Off",0,2,2,DT_QUANTITY)
'Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorQuantity\.0", "", 0, DT_QUANTITY)

Call ClickWebElementFrameNoInnertext(1, "ProcurementCatalog7Ehp1", "SPAN", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorCheckBox\.0", 0)
Call CaptureWebScreen(1,"ItemSelection")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Item Addition1")
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 15
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Call CaptureWebScreen(1,"Capture screen:Account Assignment1")
Wait 5
Call ClickSAPList(1, "Account Assignment Category", "INPUT",0)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5

Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",2,DT_SAPSRMWDC_UI_SC_ASSIGN_NO_EDIT1)
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

wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)
Call CaptureWebScreen(1,"Click Add Item2")
Call ClickWebElementSAPFrame(1, "Floor Plan Manager application for OIF", ".*text", "DLL SRM-MDM Pre\. EU - BE - ALL", "lsLink.*", 0)

Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMCJH\.BasicView\.KeywordCriteriaInputField", "text", 0, DT_KEYWORD_OCC1)
Call CaptureWebScreen(1,"Enter Search KeyWord")

Call ClickFrameSAPButton(1,"ProcurementCatalog7Ehp1","Search","DIV",0)
Call CaptureWebScreen(1,"SearchResult1")
'Call SetWebEditFrame(1, "ProcurementCatalog7Ehp1", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorQuantity\.0", "text", 0, DT_QUANTITY_OCC1)
Call SetCellDataSAPTableFrame(1,"ProcurementCatalog7Ehp1","urST3BdBrd urST3Bd urFontStd","TABLE","Add to CartAdd to Shopping ListCompareResults ViewResults Per PageImage Off",0,2,2,DT_QUANTITY_OCC1)

Call ClickWebElementFrameNoInnertext(1, "ProcurementCatalog7Ehp1", "SPAN", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorCheckBox\.0", 0)
Call CaptureWebScreen(1,"ItemSelection2")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Item Addition2")
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)

Call ClickFrameSAPButton(1,"Shopping Cart","Next Item","IMG",0)
WAit 10
Call CaptureWebScreen(1,"Capture screen:Account Assignment2")

Call ClickSAPList(1, "Account Assignment Category", "INPUT",0)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5

Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",2,DT_SAPSRMWDC_UI_SC_ASSIGN_NO_EDIT2)
Wait(1)
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

''Call GetValueWebElementFrame(1, "Display Document:", "WD1258", "SPAN", "ls.*", "DT__OCC2_OUTPUT")
''Prior to code update
'Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_LIST_TEXTVIEW\.2", "SPAN", "ls.*", "DT__OCC2_OUTPUT")
'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", DT__OCC2_OUTPUT, "ls.*", 0, False)
''Post code update
Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_2","DIV", "ls.*", "DT__OCC2_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT__OCC2_OUTPUT, "ls.*", 0, False)
Wait 5
Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
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
wait 10
Browser("CreationTime:=0").Maximize

Call LogoffSRM(0)
Call FinalStatus()


'
''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

