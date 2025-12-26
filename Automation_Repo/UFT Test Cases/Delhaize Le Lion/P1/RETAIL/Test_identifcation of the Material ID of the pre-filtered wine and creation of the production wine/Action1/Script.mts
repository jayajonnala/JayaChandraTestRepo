
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_identifcation of the Material ID of the pre-filtered wine and creation of the production wine
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

gstrTestCaseName = "TC_06_Test_identifying the Material ID and creation of the production wine"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''----------------------Tcode SE16N----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Table","GD-TAB", "", DT_TABLE_NAME, False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", "", "Fld name", "Component", DT_ARTICLE, False)
Call TakeScreenShot()
Call ClickButton("Online   \(F8\)",False)
Call TakeScreenShot()

Call GetGridContent("", "", "Bill of material", 1, "Component", DT_ARTICLE, "DT_BOM_VALUE_OUTPUT")
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call SetTextbox("Table","GD-TAB", "", DT_TABLE_NAME_OCC1, False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", "", "Fld name", "BOM", DT_BOM_VALUE_OUTPUT, False)
Call TakeScreenShot()
Call ClickButton("Online   \(F8\)",False)
Call TakeScreenShot()

Call GetGridContent("", 0, "Article", 1, "Bill of material", DT_BOM_VALUE_OUTPUT, "DT_HEADER_ARTICLE_OUTPUT")
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

''----------------------Tcode CS03----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Article","RC29N-MATNR", "", DT_HEADER_ARTICLE_OUTPUT, False)
Call SetTextbox("Site","RC29N-WERKS", "", DT_SITE, False)
Call SetTextbox("BOM Usage","RC29N-STLAN", "", DT_BOM_USAGE, False)
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
Call TakeScreenShot()

Call VerifyTableCellContent(1, "Component", "SAPLCSDITCMAT", DT_ARTICLE)

''----------------------Tcode CO01----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE_OCC2)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Article","CAUFVD-MATNR", "", DT_HEADER_ARTICLE_OUTPUT, False)
Call SetTextbox("Production Site","CAUFVD-WERKS", "", DT_PRODUCTION_SITE, False)

Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Total Qty","CAUFVD-GAMNG", "", DT_TOTAL_QUANTITY, False)
Call TakeScreenShot()
Call SetCombo("Type", "Current date")
Call TakeScreenShot()

Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'If  VerifyStatusBarConditon("is not a working day") = True Then
'	Call PressEnter() 
'	Wait(2)
'       Call TakeScreenShot()
'End If

Call ClickButtonifexist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")

Call ClickBUtton("Component Overview   \(F6\)",False)
Call TakeScreenShot()

Call ClickButtonIfexist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call ClickBUtton("Select All",False)
Call TakeScreenShot()

Call ClickButton("Batch determ\.",False)
Call TakeScreenShot()

Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call SetTextbox("Split quantity","XV01FDP-MENGE_P", "", DT_TOTAL_QUANTITY, False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F5\)",False)
Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F5\)",False)
Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F5\)",False)
Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F5\)",False)
Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call ClickButtonIfExist("Copy   \(F5\)",False)
Call ClickBUttonIfexist("Cont\.",True)
Call TakeScreenShot()

Call SetTableData("SAPLCOMKTCTRL_0120", "Reqmt Qty", 1, "", "", "0", False)
Call TakeScreenShot()
Call ClickButton("Operation Overview   \(F5\)",False)
Call TakeScreenShot()
Call ClickButtonifexist("Yes",True)
Call TakeScreenShot()

Call SetTableData("SAPLCOVGTCTRL_0100", "Work Center", 1, "", "", DT_WORK_CENTER, False)
Call TakeScreenShot()
Call Pressenter()
Call TakeScreenShot()

Call ClickButton("Release Order   \(Ctrl\+F1\)",False)
Call TakeScreenShot()

Call VerifyTextBoxContent("Error Messages","G_ERRORS", "", "0", True)
Call ClickButton("Continue Release/Derivation   \(Enter\)",True)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_ORDER_NUMBER_OUTPUT")
Call VerifyStatusBar("Order number "&DT_ORDER_NUMBER_OUTPUT&" saved")

Call LogOff()
Call FinalStatus()
