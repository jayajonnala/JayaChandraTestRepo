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

'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Maintain Asset Master Data_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Maintain Asset Master Data_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Maintain Asset Master Data_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_1140_INVENTORY_NUMBER",(Cint(DT_AS01_1140_INVENTORY_NUMBER)+3))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_EQUIPMENT",(Cint(DT_INCREMENT_EQUIPMENT)+3))
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

Call ClickButton("Master data   \(F7\)",False) 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Asset main no\. text","ANLH-ANLHTXT","",DT_AS01_1140_ASSET_MAIN_NO_TEXT,False)
Call SetTextbox("Inventory number","ANLA-INVNR","",DT_AS01_1140_INVENTORY_NUMBER,False)
Call PressEnter() 
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

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call SetTextbox("WBS element","ANLA-POSNR","",DT_WBS_ELEMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Maintain",True)
wait(3)
'Capture the screenshot
Call TakeScreenShot()

Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Description",2,DT_AS01_2300_TABLECELL_DESCRIPTION_1,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Description",3,DT_AS01_2300_TABLECELL_DESCRIPTION_2,False)

Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Inventory no.",2,DT_AS01_2300_TABLECELL_INVENTORY_NO_1,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Inventory no.",3,DT_AS01_2300_TABLECELL_INVENTORY_NO_2,False)

Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Cost Center",1,DT_AS01_2300_TABLECELL_COST_CENTER_0,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Cost Center",2,DT_AS01_2300_TABLECELL_COST_CENTER_1,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Cost Center",3,DT_AS01_2300_TABLECELL_COST_CENTER_2,False)

Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Location",1,DT_AS01_2300_TABLECELL_LOCATION_0,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Location",2,DT_AS01_2300_TABLECELL_LOCATION_1,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Location",3,DT_AS01_2300_TABLECELL_LOCATION_2,False)

Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Profit Center",1,DT_AS01_2300_TABLECELL_PROFIT_CENTER_0,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Profit Center",2,DT_AS01_2300_TABLECELL_PROFIT_CENTER_1,False)
Call SetTableDataNoRef("SAPLAISTTC_MULTIPLE_ASSETS","Profit Center",3,DT_AS01_2300_TABLECELL_PROFIT_CENTER_2,False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item2","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item3","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

