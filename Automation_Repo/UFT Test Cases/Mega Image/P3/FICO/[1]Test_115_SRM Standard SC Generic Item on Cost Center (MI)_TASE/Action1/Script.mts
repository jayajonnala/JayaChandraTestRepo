
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_115_SRM Standard SC Generic Item on Cost Center (MI)

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

gstrTestCaseName = "Test_115_SRM Standard SC Generic Item on Cost Center (MI)_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\TASEWork\Data\TASE_DT_114_SRM Standard SC Free Text on Cost Cent P1_.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''''''Close All Browser
Call CloseAllBrowsers()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUMBER",(Cint(DT_NUMBER)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''launch adn Login SRM Application
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
wait (10)
Call CaptureWebScreen(0,"Capture Home Screen")

Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Browser("CreationTime:=1").Fullscreen
Wait 30


''Call SetWebEditFrame(1, "Create Shopping Cart", "WDA6", "text", 0, DT_NAME_OF_SHOPPING_CART)
Call SetWebEditFrame(1, "Create Shopping Cart", "/SAPSRM/WDC_UI_SC_DOFC_HD\.ID_77C46A3FE491AE899896FA457CF4DA7F:V_SC_BOFC_HD\.DESCRIPTION", "text", 0, DT_NAME_OF_SHOPPING_CART)

Wait(5)
Call CaptureWebScreen(1,"Shopping Cart Description")

Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)
Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(30)
Call CaptureWebScreen(1,"Click Add Item")
''Call ClickWebElementSAPFrame(1,"Floor Plan Manager application for OIF","WD0461-text","MI SRM-MDM Pre\. EU - RO","lsLink__text",0)
Call ClickWebElementSAPFrame(1,"Floor Plan Manager application for OIF","/SAPSRM/WDC_SH_CATALOGS\.ID_FEBF028F729D3D320AF9860321669ADF:V_CATALOGS\.CI000010102-text","MI SRM-MDM Pre\. EU - RO","lsLink__text",0)
Wait 15
Call CaptureWebScreen(1,"Click on The link")
Call ClickWebElementFrame(1, "ProcurementCatalog7Ehp1", "SPAN", "Advanced Search", 0)
Wait 10
Call CaptureWebScreen(1,"Click on The link")

CAll SetWebEditFrameByTitle(1, "ProcurementCatalog7Ehp1","Enter a value for Supplier Part Number","text", "", DT_POKDMCJHADVANCEDVIEWINPUT4)
Call CaptureWebScreen(1,"Click on The link")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Search", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Search")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Not selected.*", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Check box selection")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Add to cart")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Wait 30
Call CaptureWebScreen(1,"Capture screen:Check out")
Wait 15
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5

''Call SetWebEditFrame(1, "Shopping Cart", "WD096C", "text", 0, DT_ORDER_QUANTITY__UNIT)
Call SetWebEditFrame(1, "Shopping Cart", "/SAPSRM/WDC_DODC_SC_I_BD\.ID_C3B4C372BB7F609EE2C58F327F2D2036:V_DODC_SC_I_BD\.QUANTITY", "text", 0, DT_ORDER_QUANTITY__UNIT)
''Call SetWebEditFrame(1, "Shopping Cart", "WD097F", "text", 0, DT_PRICE__CURRENCY)
Call SetWebEditFrame(1, "Shopping Cart", "/SAPSRM/WDC_DODC_SC_I_BD\.ID_C3B4C372BB7F609EE2C58F327F2D2036:V_DODC_SC_I_BD\.PRICE", "text", 0, DT_PRICE__CURRENCY)
Call CaptureWebScreen(1,"Quantity updation")


Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Wait 5

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing
Call CaptureWebScreen(1,"Details Screen")
''Call SetWebEditFrame(1, "Shopping Cart", "WD0B07", "text", 0, DT_COST_CENTER)
Call SetWebEditFrame(1, "Shopping Cart", "/SAPSRM/WDC_UI_DO_ACC\.ID_A56B0481584E314F46F1C5E908E8819E:V_DO_ACCOUNT_DETAIL\.COST_CENTER", "text", 0, DT_COST_CENTER)
Call CaptureWebScreen(1,"Cost Center updation")

Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait 15
Call CaptureWebScreen(1,"Capture screen:Complete order")

Call ClickFrameSAPButton(1,"Shopping Cart","Order","DIV",0)
Wait 30
Call CaptureWebScreen(1,"Capture screen:Complete order")
''Call GetValueWebElementFrame(1, "Display Document:", "WD0BEE", "SPAN", "ls.*", "DT_CREATED_CART_OUTPUT")
'Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_LIST_TEXTVIEW\.2", "SPAN", "ls.*", "DT_CREATED_CART_OUTPUT")
'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)

Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_2", "DIV", "ls.*", "DT_CREATED_CART_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)

Wait 15
Call ClickFrameSAPButton(1,"Display Document:","Refresh","DIV",0)
Call CaptureWebScreen(1,"Capture screen:Refresh")

Call ClickWebElementFrame(1, "Document:", "SPAN", "Display / Edit Agents", 0)
Wait 45
Call CaptureWebScreen(1,"Capture screen:Approval Flow")
Call ClickSAPFrameSAPButton(1, "Floor Plan Manager application for OIF", "OK", "DIV", 0)
Wait 10
Call ClickFrameSAPButton(1,"Document:","Close","DIV",0)
WAit 15
Call CaptureWebScreen(0,"Capture screen:Close button")
Browser("CreationTime:=0").Maximize
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet) 

Call LogoffSRM(0)
Call FinalStatus()


''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet Call

'**************************************************************************************************************************

