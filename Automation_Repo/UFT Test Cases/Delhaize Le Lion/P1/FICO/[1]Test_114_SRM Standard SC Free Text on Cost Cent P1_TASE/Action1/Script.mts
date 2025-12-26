
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_114_SRM Standard SC Free Text on Cost Cent P1_TASE
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

gstrTestCaseName = "Test_114_SRM Standard SC Free Text on Cost Cent P1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Close All Browser
Call CloseAllBrowsers()

''''launch and Login SRM Application
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(20)
Call CaptureWebScreen(0,"URL_Launch")
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)


wait (5)
''''''Click on Create Professional Shopping Cart
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call CaptureWebScreen(0,"Capture Home Screen")
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Browser("CreationTime:=1").Fullscreen
Call CaptureWebScreen(1,"Shopping Cart")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUMBER",(Cint(DT_NUMBER)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetSAPEdit(1, "Name of shopping cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
''''Enter the Description
Wait(1)
Call CaptureWebScreen(1,"Shopping Cart Description")

Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)
Call CaptureWebScreen(1,"Click Add Item")

Set wsh = createobject("Wscript.Shell")
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)

Call CaptureWebScreen(1,"From Catalog")
'Call SetWebEditSAPFrameByTitle(1, "Floor Plan Manager application for OIF", "Item Description", "text", 0, DT_SAPSRM_UI_SC_DESC_EDIT_E1_OCC1)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Description", "text", 0, DT_SAPSRM_UI_SC_DESC_EDIT_E1_OCC1)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Product Category", "text", 0, DT_SAPSRM_UI_SC_CATEGORY_ID_EDITOR11)
CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Quantity / Unit", "text", 0, DT_SAPSRM_UI_SC_QUANTITY_EDITOR1)
CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Unit for Order Quantity", "text", 0, DT_SAPSRM_UNIT_EDITOR11)
CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Net Price / Currency", "text", 0, DT_SAPSRM_PRICE_EDITOR1)
Call CaptureWebScreen(1,"Add Item With Free Description")
Call ClickSAPFrameSAPButton(1,  "Floor Plan Manager application for OIF", "Add to Item Overview", "DIV", 0)
Wait 5
Call CaptureWebScreen(1,"item Details")

Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5
'
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing
Call CaptureWebScreen(1,"item Details_AccountDetails1")

'Call SelectValSAPListAttachedText(1, "Shopping Cart", "Account Assignment Category", "INPUT", " Cost Center")

 
Call ClickSAPList(1, "Account Assignment Category", "INPUT",0)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)

Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",0,DT_COST_CENTER)

Call CaptureWebScreen(1,"CostCenter1")
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Sources of Supply / Service Agents", 0)
Wait 10
''''Call SetWebEditFrame(1, "Shopping Cart", "WD0678", "text", 0, DT_SUPPLIER)
CAll SetWebEditFrameByTitle(1, "Shopping Cart", "Fixed Supplier", "text", 0, DT_SUPPLIER)
Wait 2
Call CaptureWebScreen(1,"Supplier1")
Call ClickFrameSAPButton(1, "Shopping Cart", "Assign Supplier", "DIV", 0)
Wait 10
Call CaptureWebScreen(1,"Supplier Details")
'
Call ClickFrameSAPButton(1,"Shopping Cart","Add Item","DIV",0)

Call CaptureWebScreen(1,"Click Add Item")
Set wsh = createobject("Wscript.Shell")
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)

CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Description", "text", 0, DT_SAPSRM_DESCRIPTION_EDIT_E2)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Product Category", "text", 0, DT_SAPSRM_CATEGORY_ID_EDITOR21)
CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Quantity / Unit", "text", 0, DT_SAPSRM_QUANTITY_EDITOR2)
CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Unit for Order Quantity", "text", 0, DT_SAPSRM_UNIT_EDITOR21)
CAll SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Net Price / Currency", "text", 0, DT_SAPSRM_PRICE_EDITOR2)
Call CaptureWebScreen(1,"Add Item With Free Description")
Call ClickSAPFrameSAPButton(1,  "Floor Plan Manager application for OIF", "Add to Item Overview", "DIV", 0)
Wait 5
Call CaptureWebScreen(1,"item Details")

Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5
Call ClickFrameSAPButton(1,"Shopping Cart","Next Item","IMG",0)
Wait 5

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing
Call CaptureWebScreen(1,"item Details_AccountDetails2")

'CAll SelectValSAPListFrame(1,"Shopping Cart","Account Assignment Category", "INPUT",0,"Cost Center")
Call ClickSAPList(1, "Account Assignment Category", "INPUT",0)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5
Call SetWebEditFrameLogicalName(1,"Shopping Cart","Cost Center","text",0,DT_COST_CENTER)

Call CaptureWebScreen(1,"CostCenter2")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Sources of Supply / Service Agents", 0)
Wait 10
CAll SetWebEditFrameByTitle(1, "Shopping Cart", "Fixed Supplier", "text", 0, DT_SUPPLIER)
Wait 2

Call ClickFrameSAPButton(1,"Shopping Cart","Assign Supplier","DIV",0)
Wait(5)
Call CaptureWebScreen(1,"Assign Supplier")

Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Check Out")

Call ClickFrameSAPButton(1,"Shopping Cart","Order","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Complete order")
'Code prior to upgrade
'Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_LIST_TEXTVIEW\.3", "SPAN", "ls.*", "DT_CREATED_CART_OUTPUT")
'Code after to upgrade
Call GetValueWebElementFrame(1, "Display Document:","WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_3","DIV","ls.*", "DT_CREATED_CART_OUTPUT")
Wait 10
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)
'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)
Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 10
Call CaptureWebScreen(1,"Capture screen:Refresh")

Call ClickWebElementFrame(1, "Document:", "SPAN", "Display / Edit Agents", 0)
Wait 60
Call CaptureWebScreen(1,"Capture screen:Approval Flow")

Call ClickSAPFrameSAPButton(1, "Floor Plan Manager application for OIF", "OK", "DIV", 0)
Wait 10

Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Wait 15
Call CaptureWebScreen(0,"Capture screen:Close button")
Browser("CreationTime:=0").Maximize

'Browser("CreationTime:=1").Close
Call LogoffSRM(0)
Call FinalStatus()


'
''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************









'Browser("Create Professional Shopping").Page("Create Professional Shopping").Frame("isolatedWorkArea").SAPList("Account Assignment Category").Select "Cost Center"
'
'Browser("Create Professional Shopping").Page("Create Professional Shopping").Frame("isolatedWorkArea_2").SAPList("Account Assignment Category").Select "Asset"
'
'Browser("Create Professional Shopping").Page("Create Professional Shopping").Frame("isolatedWorkArea_2").SAPList("Account Assignment Category").Type "Asset"
'
'AIUtil.SetContext Browser("creationtime:=1")
'AIUtil("combobox", "Account Assignment Category:").Select "Cost Center"
'
'
'Browser("Create Professional Shopping").Page("Create Professional Shopping").SAPFrame("Floor Plan Manager application").SAPList("Period Type").Select "On"
'
'' ClickSAPListUniversal(creationTime, sapattachedText)
'Call ClickSAPListUniversal(1,"Period Type")
'Call ClickWebElementSAPFrame(1,"Floor Plan Manager application for OIF","WDWL1-cntTD","DescriptionProduct.*","urPWContentTableCell",0)







