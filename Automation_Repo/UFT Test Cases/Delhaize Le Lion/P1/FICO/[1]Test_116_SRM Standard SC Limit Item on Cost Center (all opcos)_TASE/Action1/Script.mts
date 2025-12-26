
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_116_SRM Standard SC Limit Item on Cost Center (all opcos)
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

gstrTestCaseName = "Test_116_SRM Standard SC Limit Item on Cost Center (all opcos)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''''Close All Browser
Call CloseAllBrowsers()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''launch and Login SRM Application
Call  LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")
wait (2)

'''Click on Create Professional Shopping Cart
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call CaptureWebScreen(0,"Capture Home Screen")
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Browser("CreationTime:=1").Fullscreen
wait (5)
Call CaptureWebScreen(1,"HomePage")
''''Call SetWebEditFrame(1, "Create Shopping Cart", "WDA6", "text", 0, DT_NAME_OF_SHOPPING_CART)
Call SetSAPEdit(1, "Name of shopping cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
Wait(2)
Call CaptureWebScreen(1,"Shopping Cart Description")

Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)
Call CaptureWebScreen(1,"Click Add Item")
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",2
Set wsh= nothing
Wait(5)
Call CaptureWebScreen(1,"Click Add Item")


Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Description", "text", 0, DT_DESCRIPTION_)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Product Category ID", "text", 0, DT_PRODUCT_CATEGORY_ID)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Value Limit", "text", 0, DT_VALUE_LIMIT)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Expected Value", "text", 0, DT_EXPECTED_VALUE)
Call CaptureWebScreen(1,"Capture screen:Complete order")

Call ClickSAPListUniversal(1,"Period Type")
'Call ClickFrameSAPList(1,"Floor Plan Manager application for OIF", "Period Type", "INPUT", 0)
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",2
Set wsh= nothing
Wait(5)
'Call SelectValSAPListLogicalName(1,"Floor Plan Manager application for OIF","Period Type", "INPUT",0,"On")
Wait 2
Call CaptureWebScreen(1,"item Details")


Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Supplier", "text", 0, DT_SUPPLIER)
Call CaptureWebScreen(1,"Add Item With Free Description")

Call ClickSAPButtonSAPFrameUniversal(1,"Floor Plan Manager application for OIF","OK","")
Wait 5
Call CaptureWebScreen(1,"Capture screen:Complete order")

Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing
Call CaptureWebScreen(1,"Account Assignment")
'Call SelectValSAPList(1,"Account Assignment Category", "INPUT",0,"Cost Center")
Call ClickSAPListUniversal(1,"Account Assignment Category")
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",2
Set wsh= nothing
Wait(5)
Call SetWebEditFrameLogicalName(1, "Shopping Cart", "Cost Center", "text", 2, DT_COST_CENTER)
Call CaptureWebScreen(1,"Cost Center Change")
Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Check1")

Call ClickFrameSAPButton(1,"Shopping Cart","Add Item","DIV",0)
Call CaptureWebScreen(1,"Click Add Item")
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)


Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Description", "text", 0, DT_DESCRIPTION__OCC1)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Product Category ID", "text", 0, DT_PRODUCT_CATEGORY_ID_OCC1)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Value Limit", "text", 0, DT_VALUE_LIMIT_OCC1)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Expected Value", "text", 0, DT_EXPECTED_VALUE)
Call CaptureWebScreen(1,"Capture screen:Complete order")
'Call SelectValSAPList(1,"Required:", "INPUT",0,"On")
Call ClickSAPListUniversal(1,"Period Type")
'Call ClickFrameSAPList(1,"Floor Plan Manager application for OIF", "Period Type", "INPUT", 0)
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",2
Set wsh= nothing
Wait(5)

Wait 2
Call CaptureWebScreen(1,"item Details")

Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Supplier", "text", 0, DT_SUPPLIER_OCC1)
Call CaptureWebScreen(1,"Add Item With Free Description")
Call ClickSAPButtonSAPFrameUniversal(1,"Floor Plan Manager application for OIF","OK","")
Wait 5
Call CaptureWebScreen(1,"Capture screen:Complete order")

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
Call CaptureWebScreen(1,"Account Assignment1")

Call ClickSAPListUniversal(1,"Account Assignment Category")
Wait 5
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",2
Set wsh= nothing
Wait(5)
Wait 5
Call SetWebEditFrameLogicalName(1, "Shopping Cart", "Cost Center", "text", 2, DT_COST_CENTER)
Wait 15

Call CaptureWebScreen(1,"Cost Center Change")

Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait 2
Call CaptureWebScreen(1,"Capture screen:Check2")

Call ClickFrameSAPButton(1,"Shopping Cart","Order","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Capture screen:Complete order")

'Code prior to upgrade
'Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_LIST_TEXTVIEW\.3", "SPAN", "ls.*", "DT_MESSAGE_OUTPUT")
'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", DT_MESSAGE_OUTPUT, "ls.*", 0, False)

'Codep Post  upgrade
Call GetValueWebElementFrame(1, "Display Document:","WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_3","DIV", "ls.*", "DT_MESSAGE_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT_MESSAGE_OUTPUT, "ls.*", 0, False)

Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Wait 5
Call CaptureWebScreen(1,"Capture screen:Refresh")

Call ClickWebElementFrame(1, "Document:", "SPAN", "Display / Edit Agents", 0)
Wait 60
Call CaptureWebScreen(1,"Capture screen:Approval Flow")

Call ClickSAPFrameSAPButton(1, "Floor Plan Manager application for OIF", "OK", "DIV", 0)
Wait 20
Call CaptureWebScreen(1,"Capture screen:After Clicking OK")

'Call ClickFrameSAPButton(1,"Display Document:","Close","DIV",0)
Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
Call CaptureWebScreen(0,"Capture screen:Close")
Browser("CreationTime:=0").Maximize
Call LogoffSRM(0)
Call FinalStatus()


'
''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

