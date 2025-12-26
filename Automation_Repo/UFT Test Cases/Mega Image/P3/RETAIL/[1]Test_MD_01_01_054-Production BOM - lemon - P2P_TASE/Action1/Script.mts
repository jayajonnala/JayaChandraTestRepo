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

''Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_MD_01_01_054-Production BOM - lemon - P2P_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 18th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_MD_01_01_054-Production  P2P_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_MD_01_01_054-Production BOM - lemon - P2P_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
''Increment the parameter/reload
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode SE16N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Table","GD-TAB","",DT_SE16N_0100_TABLE,False)
Call PressEnter()
Call TakeScreenShot

Call ClickButtonIfExist("Deselect All   \(Shift\+F6\)",False)

Call ClickButtonIfExist("Find   \(Ctrl\+F\)",False)
Call SetTextbox("Field Name","SVALD-VALUE","",DT_SE16N_0300_FIELD_NAME,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTableData("SAPLSE16NSELFIELDS_TC","Output","","Technical name",DT_SE16N_0300_FIELD_NAME,DT_SE16N_0100_TABLECELL_OUTPUT_0,False)
Call TakeScreenShot

Call ClickButtonIfExist("Find   \(Ctrl\+F\)",False)
Call SetTextbox("Field Name","SVALD-VALUE","",DT_SE16N_0300_FIELD_NAME_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTableData("SAPLSE16NSELFIELDS_TC","Output","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC1,DT_SE16N_0100_TABLECELL_OUTPUT_0_OCC1,False)
Call SetTableData("SAPLSE16NSELFIELDS_TC","Fr.Value","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC1,DT_SE16N_0100_TABLECELL_FRVALUE_0,False)
Call TakeScreenShot

Call ClickButtonIfExist("Find   \(Ctrl\+F\)",False)
Call SetTextbox("Field Name","SVALD-VALUE","",DT_SE16N_0300_FIELD_NAME_OCC2,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SelectCellGuiTable("SAPLSE16NSELFIELDS_TC","Option","Technical name",DT_SE16N_0300_FIELD_NAME_OCC2,False)
Call PressEnter()

Call SelectCellGuiGrid("","",DT_SE16N_0500_GRIDCELL_3_TEXT,"Text",True)
Call SelectRowGuiGrid("","","Text",DT_SE16N_0841_SEARCH_TERM,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Online   \(F8\)",False)
Call TakeScreenShot

Call GetGridContentByTitle("","","Article",1,"DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Find   \(Ctrl\+F\)",False)
Call SetTextbox("Field Name","SVALD-VALUE","",DT_SE16N_0300_FIELD_NAME_OCC3,True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTableData("SAPLSE16NSELFIELDS_TC","Fr.Value","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC3,DT_SE16N_0100_TABLECELL_FRVALUE_0_OCC1,False)
Call TakeScreenShot

Call ClickButtonIfExist("Online   \(F8\)",False)
Call TakeScreenShot

Call GetGridContentByTitle("","","Article",1,"DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OCC1_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode WSM3----------------------------
'Enter the Tcode
Call SetTcode(DT_SE16N_0200_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SE16N_0200_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call SetTextbox("Variant","V-LOW","",DT_SE16N_0841_SEARCH_TERM_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",True)

SAPGuiSession("transaction:=WSM3").SAPGuiWindow("transaction:=WSM3").SAPGuiButton("tooltip:=Multiple selection","index:=1").Click
Wait(2)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_SE16N_0100_ARTICLE,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_SE16N_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",3,DT_SE16N_3010_TABLECELL_SINGLE_VALUE_2,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyStatusBar(DT_SE16N_0120_CHECK_TEXT_OF_STATUSBAR)
''
''''----------------------Tcode CS01----------------------------
'Enter the Tcode
Call SetTcode(DT_SE16N_0120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SE16N_0120_OKCD)

Call SetTextbox("Article","RC29N-MATNR","",DT_SE16N_0100_ARTICLE,False)
Call SetTextbox("Site","RC29N-WERKS","",DT_SE16N_0100_SITE,False)
Call SetTextbox("BOM Usage","RC29N-STLAN","",DT_SE16N_0100_BOM_USAGE,False)
'Capture the screenshot
Call TakeScreenShot()

''Click button Check entries
Call ClickButtonIfExist("Check entries   \(Shift\+F5\)",False)
Call TakeScreenShot

Call PressEnter()
Wait(5)
Call PressEnter()  
Wait(2)
Call SetTableDataNoRef("SAPLCSDITCMAT","ICt",1,DT_SE16N_0152_TABLECELL_ICT_0,False)
Call SetTableDataNoRef("SAPLCSDITCMAT","ICt",2,DT_SE16N_0152_TABLECELL_ICT_1,False)

Call SetTableDataNoRef("SAPLCSDITCMAT","Component",1,DT_SE16N_0152_TABLECELL_COMPONENT_0,False)
Call SetTableDataNoRef("SAPLCSDITCMAT","Component",2,DT_SE16N_0152_TABLECELL_COMPONENT_1,False)

Call SetTableDataNoRef("SAPLCSDITCMAT","Quantity",1,DT_SE16N_0152_TABLECELL_QUANTITY_0,False)
Call SetTableDataNoRef("SAPLCSDITCMAT","Quantity",2,DT_SE16N_0152_TABLECELL_QUANTITY_1,False)

Call SetTableDataNoRef("SAPLCSDITCMAT","Un",1,DT_SE16N_0152_TABLECELL_UN_0,False)
Call SetTableDataNoRef("SAPLCSDITCMAT","Un",2,DT_SE16N_0152_TABLECELL_UN_1,False)
Call PressEnter()  
Call TakeScreenShot

'' Click on Save
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
wait(2)
Call TakeScreenShot
Call VerifyStatusBar(DT_SE16N_0100_CHECK_TEXT_OF_STATUSBAR)

''
''''----------------------Tcode CS03----------------------------
'Enter the Tcode
Call SetTcode(DT_SE16N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SE16N_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","RC29N-MATNR","",DT_SE16N_0100_ARTICLE_OCC1,False)
Call SetTextbox("Site","RC29N-WERKS","",DT_SE16N_0100_SITE_OCC1,False)
Call SetTextbox("BOM Usage","RC29N-STLAN","",DT_SE16N_0100_BOM_USAGE_OCC1,False)
Call SetTextbox("Valid From","RC29N-DATUV","",DT_SE16N_0100_VALID_FROM,False)
Call SetTextbox("Valid to","RC29N-DATUB","",DT_SE16N_0100_VALID_TO,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Check entries   \(Shift\+F5\)",False)
Call TakeScreenShot()
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

'Call SelectRowGuiTable("SAPLCSDITCALT","AltBOM",1,False)
'Call ClickButtonIfExist("Item Overview   \(F5\)",False)
'Call TakeScreenShot()

Call VerifyTableCellContent(1,"ICt","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_ICT_0)
Call VerifyTableCellContent(2,"ICt","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_ICT_1)

Call VerifyTableCellContent(1,"Component","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_0)
Call VerifyTableCellContent(2,"Component","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_1)

Call VerifyTableCellContent(1,"Quantity","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_QUANTITY_0)
Call VerifyTableCellContent(2,"Quantity","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_QUANTITY_1)

Call VerifyTableCellContent(1,"Un","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_UN_0)
Call VerifyTableCellContent(2,"Un","SAPLCSDITCMAT",DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_UN_1)


Wait(2)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

