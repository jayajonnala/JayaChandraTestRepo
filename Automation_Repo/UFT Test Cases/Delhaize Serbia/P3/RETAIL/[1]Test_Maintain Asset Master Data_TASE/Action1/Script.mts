'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Maintain Asset Master Data_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Maintain Asset Master Data_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Maintain Asset Master Data_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_EQUIPMENT",(Cint(DT_INCREMENT_EQUIPMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_DESCRIPTION",(Cint(DT_INCREMENT_DESCRIPTION)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS01----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Asset main no\. text","ANLH-ANLHTXT","",DT_AS01_1140_ASSET_MAIN_NO_TEXT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Location","ANLZ-STORT","",DT_AS01_1145_LOCATION,False)

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Deprec. Areas",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'
'_________________________Iter 2__________________________
'
Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE_OCC1,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Master data   \(F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION_OCC1,False)
Call SetTextbox("Asset main no\. text","ANLH-ANLHTXT","",DT_AS01_1140_ASSET_MAIN_NO_TEXT_OCC1,False)
Call SetTextbox("Inventory number","ANLA-INVNR","",DT_AS01_1140_INVENTORY_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER_OCC1,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2_OCC1,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Deprec. Areas",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCellGuiTable("SAPLAISTTC_ANLB","DKey","Area number","30",False)'substitue of selecting cell of row no 5
Call SendKey("{F4}")
'Capture the screenshot
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)

Call ClickButton("Cancel   \(F12\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1_OCC2,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call SetTextbox("WBS element","ANLA-POSNR","",DT_WBS_ELEMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
''''''''''''''''''''''''''''''''''''''''''''''''''''
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
''''''''''''''''''''''''''''''''''''''''''''''''''''
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

