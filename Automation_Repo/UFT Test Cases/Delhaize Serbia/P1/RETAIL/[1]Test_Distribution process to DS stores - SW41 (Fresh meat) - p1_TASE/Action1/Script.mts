
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p1
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW41 (Fresh meat) - p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VL02N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL02N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Enter the Planned GI Date
Call SetTextbox("Planned GI","LIKP-WADAT","",DT_VL02N_1102_PLANNED_GI,FALSE)
Call TakeScreenShot()

'Click on Pack Menu button
Call ClickButton("Pack   \(Shift\+F6\)",False) 
Wait(2)

''Call PressEnter()
'Call VerifyStatusBarMessageType("W")
Call PressEnter()
Wait(1)
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Set The Packaging Article No
Call SetTableDataNoRef("SAPLV51GTC_HU_001","Packaging Articles",1,DT_VL02N_6010_TABLECELL_PACKAGING_ARTICLES_0,False)
Call TakeScreenShot()

'Click on Detailed view W/Vol
Call ClickButton("Detailed view W/Vol",False) 
Wait(1)
Call TakeScreenShot()

'Navigate to Status Tab
Call SelectTab("TS_HU_DET","Status",False)
Wait(1)

'Enter site details
Call SetTextbox("Site/storage loc\.","VEKPVB-WERKS","",DT_VL02N_6150_SITESTORAGE_LOC,False)
Call SetTextbox("/","VEKPVB-LGORT","",DT_VL02N_6150_VEKPVBLGORT,False)
Call TakeScreenShot()

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'verify newly added Item
Call VerifyTableCellContent(5,"Itm","SAPMV50ATC_LIPS_OVER",DT_VL02N_1102_CHECK_TEXT_OF_TABLECELL_ITM_4)

'Click on Save
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(1)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(1)
Call TakeScreenShot()

''Click on Continue
'Call ClickButtonIfExist("Continue   \(Enter\)",False)
'wait(1)
'Call ClickButtonIfExist("Continue   \(Enter\)",False)
'wait(1)
'Call ClickButtonIfExist("Continue   \(Enter\)",False)
'

'Verify the status Bar message
Call VerifyStatusBar(DT_VL02N_4004_CHECK_TEXT_OF_STATUSBAR)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

