'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Manage Asset Transfer (intra OpCo)_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 16th March
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

gstrTestCaseName = "Test_Manage Asset Transfer (intra OpCo)_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Manage Asset Transfer (intra OpCo)_p3_TASE.xls"
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
'''----------------------Tcode ABUMN----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABUMN_0300_COMPANY_CODE,True)
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABUMN_0300_ASSET,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABUMN_0206_TEXT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_ABUMN_0320_EXISTING_ASSET,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",Replace(DT_ABUMN_0202_ASSET_VALUE_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Posting Date","RAIFP1-BUDAT","",Replace(DT_ABUMN_0201_POSTING_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","RAIFP1-BLDAT","",Replace(DT_ABUMN_0200_DOCUMENT_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Additional Details",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document type","RAIFP1-BLART","",DT_ABUMN_0204_DOCUMENT_TYPE,False)
Call SetTextbox("Posting period","RAIFP2-MONAT","",DT_ABUMN_0203_POSTING_PERIOD,False)
Call SetTextbox("Transfer variant","RAIFP1-TRAVA","",DT_ABUMN_0210_TRANSFER_VARIANT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Partial Transfer",False)
'Capture the screenshot
Call TakeScreenShot()

'''Call SelectRadioButtonIfExist("RAIFP2-XANEU",2,FALSE)
Call SelectRadioButton("RAIFP2-XANEU",DT_ABUMN_0401_FROM_CURRYR_AQUIS,FALSE)
Call SetTextbox("Percentage rate","RAIFP2-PROZS","",DT_ABUMN_0401_PERCENTAGE_RATE,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Simulate   \(F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'validate grid components
Call VerifyGridCellContent("Line items",1,"Posting Key","",DT_ABUMN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("Line items",2,"Posting Key","",DT_ABUMN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("Line items",1,"G/L Account","",DT_ABUMN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)

Call ClickButton("Back   \(F3\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item2","DT_ABUMN_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ABUMN_0100_CHECK_TEXT_OF_STATUSBAR)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

