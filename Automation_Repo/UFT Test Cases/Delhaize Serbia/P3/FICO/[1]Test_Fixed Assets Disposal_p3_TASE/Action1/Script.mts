'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Fixed Assets Disposal_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 18th March
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

gstrTestCaseName = "Test_Fixed Assets Disposal_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Disposal_p5_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABAVN_0300_ASSET,False)
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
'''----------------------Tcode ABAVN----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

''''Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABAVN_0300_COMPANY_CODE,True)
''''Call SetTextboxPopupIfExist("SVALD-VALUE","Company Code",DT_ABAVN_0300_COMPANY_CODE)
''''Call PressEnter() 
'''''Capture the screenshot
''''Call TakeScreenShot()


Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABAVN_0300_ASSET,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABAVN_0300_ASSET_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",Replace(DT_ABAVN_0202_ASSET_VALUE_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Posting Date","RAIFP1-BUDAT","",Replace(DT_ABAVN_0201_POSTING_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","RAIFP1-BLDAT","",Replace(DT_ABAVN_0200_DOCUMENT_DATE,"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABAVN_0206_TEXT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Additional Details",False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document type","RAIFP1-BLART","",DT_ABAVN_0204_DOCUMENT_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Partial retirement",False)
'Capture the screenshot
Call TakeScreenShot()

'''Call SelectRadioButtonIfExist("RAIFP2-XANEU",2,FALSE)
Call SelectRadioButton("RAIFP2-XANEU",DT_ABAVN_0401_FROM_CURRYR_AQUIS,FALSE)
Call SetTextbox("Percentage rate","RAIFP2-PROZS","",DT_ABAVN_0401_PERCENTAGE_RATE,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Simulate   \(F9\)",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

'getting session details for future
Call GetGridContentByTitle("Line items","","Posting Key",1,"DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OUTPUT")
Call GetGridContentByTitle("Line items","","Posting Key",2,"DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OUTPUT")
Call GetGridContentByTitle("Line items","","G/L Account",1,"DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OUTPUT")
Call GetGridContentByTitle("Line items","","G/L Account",2,"DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT_OUTPUT")
'loading datasheet again
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'validate grid components
call VerifyGridCellContent("Line items",1,"Posting Key","",DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
call VerifyGridCellContent("Line items",2,"Posting Key","",DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
call VerifyGridCellContent("Line items",1,"G/L Account","",DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
call VerifyGridCellContent("Line items",2,"G/L Account","",DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT_OCC1)

Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("RAIFP2-BUKRS","","DT_ABAVN_0300_CHECK_TEXT_OF_COMPANY_CODE_OUTPUT",False)
'loading datasheet again
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'validate company code
Call VerifyTextBoxContent("Company Code","RAIFP2-BUKRS","",DT_ABAVN_0300_CHECK_TEXT_OF_COMPANY_CODE_OCC1,False)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item2","DT_DOCUMENT_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ABAVN_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

