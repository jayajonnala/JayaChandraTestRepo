
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_149c_SRM Standard SC Generic Item Check Budget after Completion
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

gstrTestCaseName = "Test_149c_SRM Standard SC Generic Item Check Budget after Completion_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'''''Close All Browser
Call CloseAllBrowsers()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_CART_NAME_INC",(Cint(DT_CART_NAME_INC)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet) 
''launch and Login SRM Application
'Call LaunchSAPWebApplication(DT_SAPURL)
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
wait (5)
Call CaptureWebScreen(0,"Capture Home Screen")
'Click on Create Professional Shopping Cart
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Wait 30
Browser("CreationTime:=1").Fullscreen
Call CaptureWebScreen(1,"Shopping Cart Description")
'Call SetWebEditFrame(1, "Create Shopping Cart", "WDA6", "text", 0, DT_NAME_OF_SHOPPING_CART)
Call SetSAPEdit(1, "Name of shopping cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
Wait(1)
Call CaptureWebScreen(1,"Shopping Cart Description")
Wait(5)

Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)

Set wsh = createobject("Wscript.Shell")
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
Wait(30)
Call CaptureWebScreen(1,"Shopping Cart Description")
Call ClickWebElementSAPFrame(1,"Floor Plan Manager application for OIF",".*text","MI SRM-MDM Pre\. EU - RO","lsLink__text",0)
Wait 15
Call CaptureWebScreen(1,"Shopping Cart Description")
Call ClickWebElementFrame(1, "ProcurementCatalog7Ehp1", "SPAN", "Advanced Search", 0)
Wait 10
CAll SetWebEditFrameByTitle(1, "ProcurementCatalog7Ehp1","Enter a value for Supplier Part Number","text", 0, DT_POKDMCJHADVANCEDVIEWINPUT4)

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Search", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Search")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Not selected.*", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Check box selection")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Add to cart")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Check out")
Wait 10

Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5

Call SetWebEditFrameLogicalName(1, "Shopping Cart", "Order Quantity / Unit", "text", 0, DT_QUANTITY_EDITOR1)
Call SetWebEditFrameLogicalName(1, "Shopping Cart", "Price / Currency", "text", 0, DT_PRICE_EDITOR1)

'Call SetWebEditFrame(1, "Shopping Cart", "WD096C", "text", 0, DT_QUANTITY_EDITOR1)
'Call SetWebEditFrame(1, "Shopping Cart", "WD097F", "text", 0, DT_PRICE_EDITOR1)
Call CaptureWebScreen(1,"Quantity updation")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing

'Call SetWebEditFrame(1, "Shopping Cart", "WD0B07", "text", 0, DT_COSTCENTER)
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",2,DT_COSTCENTER)
Call CaptureWebScreen(1,"Cost Center updation")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Item Data", 0)
Call CaptureWebScreen(1,"Shopping Cart Description")
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Plant / Location","text",1,7625)

Set wsh = createobject("Wscript.Shell")
Call ClickFrameSAPButton(1,"Shopping Cart","Add Item","DIV",0)
Set wsh = createobject("Wscript.Shell")
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
Wait(30)
Call CaptureWebScreen(1,"Shopping Cart Description")

'Call ClickWebElementSAPFrame(1,"Floor Plan Manager application for OIF","WD0461-text","MI SRM-MDM Pre\. EU - RO","lsLink__text",0)
Call ClickWebElementSAPFrame(1,"Floor Plan Manager application for OIF",".*text","MI SRM-MDM Pre\. EU - RO","lsLink__text",0)

Wait 15

Call CaptureWebScreen(1,"Shopping Cart Description")
Call ClickWebElementFrame(1, "ProcurementCatalog7Ehp1", "SPAN", "Advanced Search", 0)
Wait 10

CAll SetWebEditFrameByTitle(1, "ProcurementCatalog7Ehp1","Enter a value for Supplier Part Number","text", 0, DT_POKDMCJHADVANCEDVIEWINPUT4_OCC1)

Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Search", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Search")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Not selected.*", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Check box selection")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Add to cart")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Check out")
Wait 10

Call ClickFrameSAPButton(1,"Shopping Cart","Next Item","IMG",0)
Wait 5

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing


Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Item Data", 0)
Call CaptureWebScreen(1,"Item Data screen")

Call SetWebEditFrameLogicalName(1, "Shopping Cart", "Order Quantity / Unit", "text", 0, DT_ORDER_QUANTITY__UNIT)
Call SetWebEditFrameLogicalName(1, "Shopping Cart", "Price / Currency", "text", 0, DT_PRICE__CURRENCY)

'Call SetWebEditFrame(1, "Shopping Cart", "WD126E", "text", 0, DT_ORDER_QUANTITY__UNIT)
'Call SetWebEditFrame(1, "Shopping Cart", "WD1281", "text", 0, DT_PRICE__CURRENCY)
Call CaptureWebScreen(1,"Quantity updation")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Wait 5
''''Call SetWebEditFrame(1, "Shopping Cart", "WD10F6", "text", 0, DT_COST_CENTER_)
'Call SetWebEditFrame(1, "Shopping Cart", "WD13D9", "text", 0, DT_COST_CENTER_)
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",2,DT_COSTCENTER)
Call CaptureWebScreen(1,"Cost Center updation")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Item Data", 0)
Call CaptureWebScreen(1,"Shopping Cart Description")
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Plant / Location","text",1,7625)

Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Call CaptureWebScreen(1,"Capture screen:Complete order")
Wait 5
Call ClickFrameSAPButton(1,"Shopping Cart","Order","DIV",0)
Call CaptureWebScreen(1,"Capture screen:Complete order")
Wait 10
'Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_LIST_TEXTVIEW\.3", "SPAN", "ls.*", "DT_CREATED_CART_OUTPUT")
'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)

Call GetValueWebElementFrame(1, "Display Document:","WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_3","DIV","ls.*", "DT_CREATED_CART_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)

Wait 5
Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Call CaptureWebScreen(1,"Capture screen:Refresh")

Call ClickWebElementFrame(1, "Document:", "SPAN", "Display / Edit Agents", 0)
Wait 60
Call CaptureWebScreen(1,"Capture screen:Approval Flow")

'Call GetValueWebElementSAPFrame(1, "Floor Plan Manager application for OIF", "WD0C8C-text", "ls.*", "DT_PROCESSOR")
Call ClickSAPFrameSAPButton(1, "Floor Plan Manager application for OIF", "OK", "DIV", 0)
Wait 10
Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 15
Call CaptureWebScreen(1,"Capture screen:Close button")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Browser("CreationTime:=0").Maximize
Call LogoffSRM(0)
Call FinalStatus()


''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

