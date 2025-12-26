'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Fixed Assets Sales_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th March
'.................Modified By :
'.................Modified Date/Details :
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Fixed Assets Sales_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p5_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'
'''----------------------Tcode AS02----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Inventory number","ANLA-INVNR","",DT_AS02_1140_INVENTORY_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Deprec. Areas",False)
'Capture the screenshot
Call TakeScreenShot()

'Click cell of row no 3,,"06" is the area nummber value of row no 3
Call SelectCellGuiTable("SAPLAISTTC_ANLB","Area number","Area number","06",False)
Call DoubleClick()
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("ANLB-XNEGA",0,DT_AS02_3095_NEG_VALS_ALLOWED,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Click cell of row no 4,,"08" is the area nummber value of row no 4
Call SelectCellGuiTable("SAPLAISTTC_ANLB","Area number","Area number","08",False)
Call DoubleClick()
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("ANLB-XNEGA",0,DT_AS02_3095_NEG_VALS_ALLOWED_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Click cell of row no 7,,"06" is the area nummber value of row no 7
Call SelectCellGuiTable("SAPLAISTTC_ANLB","Area number","Area number","99",False)
Call DoubleClick()
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckbox("ANLB-XNEGA",0,DT_AS02_3095_NEG_VALS_ALLOWED_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
wait(1)
Call TakeScreenShot()

'Validate If asset is changed
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS02_0100_CHECK_TEXT_OF_STATUSBAR)

'
'''----------------------Tcode ABUMN----------------------------
'
'Enter the Tcode
Call SetTcode(DT_AS02_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_AS02_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",Replace(DT_AS02_0202_ASSET_VALUE_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","RAIFP1-BLDAT","",Replace(DT_AS02_0200_DOCUMENT_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_AS02_0320_EXISTING_ASSET,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Additional Details",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document type","RAIFP1-BLART","",DT_AS02_0204_DOCUMENT_TYPE,False)
Call SetTextbox("Transfer variant","RAIFP1-TRAVA","",DT_AS02_0210_TRANSFER_VARIANT,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Partial Transfer",False)
'Capture the screenshot
Call TakeScreenShot()

'''Call SelectRadioButtonIfExist("RAIFP2-XANEU",2,FALSE)
Call SelectRadioButton("RAIFP2-XANEU",DT_AS02_0401_FROM_CURRYR_AQUIS,FALSE)
Call SetTextbox("Percentage rate","RAIFP2-PROZS","",DT_AS02_0401_PERCENTAGE_RATE,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Simulate   \(F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
'
'''''Call ClickButton("Back   \(F3\)",False)
'''''Wait(1)
''''''Capture the screenshot
'''''Call TakeScreenShot()
'''''Call ClickButtonIfExist("Yes",True)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item2","DT_AS02_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS02_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

