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


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AA072 Transfer Asset Current year btw assets same CC_P2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 24th March
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_AA072 Transfer Asset Current year btw assets same CC_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_AA072 Transfer Asset Current year btw assets same CC_P2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''Login'''
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
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABUMN_0300_ASSET,False)
Call SetTextbox("Asset","RAIFP2-ANLN2","",DT_ABUMN_0300_ASSET_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABUMN_0206_TEXT,False)
Call SetTextbox("Existing asset","RAIFP3-ANLN1","",DT_ABUMN_0320_EXISTING_ASSET,False)
Call SetTextbox("Existing asset","RAIFP3-ANLN2","",DT_ABUMN_0320_EXISTING_ASSET_OCC1,False)
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
Call SetTextbox("Reference","RAIFP1-XBLNR","",DT_ABUMN_0207_REFERENCE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Partial Transfer",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButton("RAIFP2-XAALT",DT_ABUMN_0401_FROM_CURRYR_AQUIS,FALSE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
wait(1)
Call TakeScreenShot()

'Validate If asset is created
Call GetStatusBar("item1","DT_ABUMN_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item2","DT_ABUMN_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot()
VerifyStatusBar(DT_ABUMN_0100_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

