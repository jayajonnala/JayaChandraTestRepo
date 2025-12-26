
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AT_MD_01_01_057-BOM - Multipack
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


gstrTestCaseName = "Test_AT_MD_01_01_057-BOM - Multipack"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''S:\TASETestData\P3\MI\RETAIL\TASE_DT_MD_01_01_057-BOM - Multipack.xls
''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_MD_01_01_057-BOM - Multipack.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-SE16N----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


 Call SetTextbox("Table","GD-TAB", "", DT_SE16N_0100_TABLE, False)
 Call TakeScreenShot
 Call PressEnter() 

 Call ClickButtonIfExist("Deselect All   \(Shift\+F6\)", False)
 Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
 Call TakeScreenShot

 Call SetTextbox("Field Name","SVALD-VALUE", "", DT_SE16N_0300_FIELD_NAME, True)
 Call ClickButtonIfExist("Continue   \(Enter\)", True)
 Call TakeScreenShot


''Call SetTableData("SAPLSE16NSELFIELDS_TC","Output",1,"Technical name","MATNR",DT_SE16N_0100_TABLECELL_OUTPUT_0,False)

''Select checkbox in table
Call SetTableData("SAPLSE16NSELFIELDS_TC","Output","","Technical name",DT_SE16N_0300_FIELD_NAME,DT_SE16N_0100_TABLECELL_OUTPUT_0,False)
Call TakeScreenShot

''Click on Find button
 Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
 Call SetTextbox("Field Name","SVALD-VALUE", "", DT_SE16N_0300_FIELD_NAME_OCC1, True)
 Call ClickButtonIfExist("Continue   \(Enter\)", True)

''Enter value R* in Fr.Value
Call SetTableData("SAPLSE16NSELFIELDS_TC","Fr.Value","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC1,DT_SE16N_0100_TABLECELL_FRVALUE_0,False)
Call TakeScreenShot

Call ClickButtonIfExist("Online   \(F8\)", False)


wait 2
'Click on SelectAll button in guiGrid
 call SelectAllRowGuiGrid("", 0, False)
 
'' Call PressEnter() 
'' Call PressEnter()
''Call SetFocusGuiLabel("Article", 4, 1, False)
Dim WshShell
 Set WshShell = CreateObject("WScript.Shell")
  'WshShell.SendKeys "^(a)"
  WshShell.SendKeys "{TAB}"
  'WshShell.SendKeys "{TAB}"
  WshShell.SendKeys "^(a)"
  WshShell.SendKeys "^(a)"
 WshShell.SendKeys "^(c)"
 
 Call TakeScreenShot
 Call ClickButtonIfExist("Continue   \(Enter\)", True)
'' Call PressEnter() 
  Call ClickButtonIfExist("Back   \(F3\)", False)

Call SetTextbox("Table","GD-TAB", "", DT_SE16N_0100_TABLE_OCC1, False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call ClickButtonIfExist("Deselect All   \(Shift\+F6\)", False)
Call TakeScreenShot

Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
Call SetTextbox("Field Name","SVALD-VALUE", "", DT_SE16N_0300_FIELD_NAME_OCC2, True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot

'''Select first checkbox in table
Call SetTableData("SAPLSE16NSELFIELDS_TC","Output","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC2,DT_SE16N_0100_TABLECELL_OUTPUT_0_OCC1,False)
Call TakeScreenShot

Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
Call SetTextbox("Field Name","SVALD-VALUE", "", DT_SE16N_0300_FIELD_NAME_OCC3, True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)
Call TakeScreenShot

Call SetTableData("SAPLSE16NSELFIELDS_TC","Output","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC3,DT_SE16N_0100_TABLECELL_OUTPUT_0_OCC2,False)
Call SetTableData("SAPLSE16NSELFIELDS_TC","Fr.Value","","Technical name",DT_SE16N_0300_FIELD_NAME_OCC3,DT_SE16N_0100_TABLECELL_FRVALUE_0_OCC1,False)
Call TakeScreenShot

Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
Call SetTextbox("Field Name","SVALD-VALUE", "", DT_SE16N_0300_FIELD_NAME_OCC4, True)
Call ClickButtonIfExist("Continue   \(Enter\)", True)

Call FindRowNumber("SAPLSE16NSELFIELDS_TC", "Technical name", DT_SE16N_0300_FIELD_NAME_OCC4, "FIND_ROW_NUMBER_OUTPUT")

Call ClickCellTable("SAPLSE16NSELFIELDS_TC","Option",FIND_ROW_NUMBER_OUTPUT,"Technical name",DT_SE16N_0300_FIELD_NAME_OCC4, False)

Call TakeScreenShot

''Click on Option icon form first row, choose option '=' and continue(enter)
 Call SelectRowGuiGridbyRowNo("", 0, DT_SE16N_0500_GRIDCELL_3_ICON, True)

 Call ClickButtonIfExist("Continue   \(Enter\)", True)
   Call TakeScreenShot
''Click on Find button
 Call ClickButtonIfExist("Find   \(Ctrl\+F\)", False)
 Call SetTextbox("Field Name","SVALD-VALUE", "", DT_SE16N_0300_FIELD_NAME_OCC5, True)
 Call ClickButtonIfExist("Continue   \(Enter\)", True)


''Click on Push button under More column
Call FindRowNumber("SAPLSE16NSELFIELDS_TC", "Technical name", DT_SE16N_0300_FIELD_NAME_OCC5, "c")

Call ClickCellTable("SAPLSE16NSELFIELDS_TC","More",FIND_ROW_NUMBER_OUTPUT,"Technical name",DT_SE16N_0300_FIELD_NAME_OCC5, False)
''Call ClickCellTableByRowNo("SAPLSE16NSELFIELDS_TC", "More", 1, False)
Call TakeScreenShot


Call ClickButtonIfExist("Upload from Clipboard   \(Shift\+F12\)", True)

Call ClickButtonIfExist("Transfer Data   \(F8\)", True)
Call TakeScreenShot
Wait 10

Call ClickButtonIfExist("Online   \(F8\)", False)
Call TakeScreenShot
Wait 10
' ======  Get Article value of first row

Call GetGridContent("", 0, "Article", 1, "<NA>", "<NA>", "DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT")
Call GetGridContent("", 0, "Article", 2, "<NA>", "<NA>", "DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MATNR_OUTPUT")


' ======= Transaction code - /nCS01

Call SetTcode(DT_SE16N_0200_OKCD)     
Call PressEnter()     
Call TakeScreenShot
''Call CheckTCodeScreen(DT_SE16N_0200_OKCD)


Call SetTextbox("Article", "RC29N-MATNR", "", DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT, False)
Call SetTextbox("BOM Usage", "RC29N-STLAN", "", DT_SE16N_0100_BOM_USAGE, False)
''Click button Check entries
Call ClickButtonIfExist("Check entries   \(Shift\+F5\)", False)
Call TakeScreenShot

'' Click enter key 4 times
Call PressEnter()  
Call PressEnter()  
Call PressEnter()  
Wait 5
'Call GetTextboxValue("RC29N-STLAL", "", "FIND_ROW_NUMBER_OUTPUT", False)
Call PressEnter()  

''create Material Bom screen
 Call SetTableDataNoRef("SAPLCSDITCMAT", "ICt", 1, DT_SE16N_0152_TABLECELL_ICT_0, False)

''Call SetTableDataNoRef("SAPLCSDITCMAT", "Component", 1, DT_SE16N_0152_TABLECELL_COMPONENT_0, False)
Call SetTableDataNoRef("SAPLCSDITCMAT", "Component", 1, DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MATNR_OUTPUT, False)

Call SetTableDataNoRef("SAPLCSDITCMAT", "Quantity", 1, DT_SE16N_0152_TABLECELL_QUANTITY_0, False)

Call SetTableDataNoRef("SAPLCSDITCMAT", "Un", 1, DT_SE16N_0152_TABLECELL_UN_0, False)
Call PressEnter()  
Call TakeScreenShot

'' Click on Save
Call ClickButtonIfExist("Save   \(Ctrl\+S\)", False)
'Call TakeScreenShot
'' Get statusbar text
 Call GetTextStatusBar(DT_SE16N_0100_CHECK_TEXT_OF_STATUSBAR)
''Creating BOM for article 7064110
''Call GetStatusBar("item1", "DT_VF01_0101_BILLING_DOCUMENT_OUTPUT")
''Call VerifyStatusBar("Creating BOM for article "&DT_SE16N_0100_ARTICLE)
Call VerifyStatusBar("Creating BOM for article "&DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT)
Call TakeScreenShot
'' ======= Transaction code - /nCS03

Call SetTcode(DT_SE16N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Article", "RC29N-MATNR", "", DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT, False)
Call SetTextbox("BOM Usage", "RC29N-STLAN", "", DT_SE16N_0100_BOM_USAGE_OCC1, False)
'Call SetTextbox("Alternative BOM","RC29N-STLAL", "", FIND_ROW_NUMBER_OUTPUT, False)
Call PressEnter()  
Call TakeScreenShot

''Validations

Call VerifyTextBoxContent("Article", "RC29K-MATNR", "", DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OUTPUT, False)

Call VerifyTableCellContent(1, "ICt", "SAPLCSDITCMAT", DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_ICT_0)
''Call VerifyTableCellContent(1, "Component", "SAPLCSDITCMAT", DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_0)
Call VerifyTableCellContent(1, "Component", "SAPLCSDITCMAT", DT_SE16N_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MATNR_OUTPUT)
Call VerifyTableCellContent(1, "Quantity", "SAPLCSDITCMAT", DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_QUANTITY_0)
Call VerifyTableCellContent(1, "Un", "SAPLCSDITCMAT", DT_SE16N_0152_CHECK_TEXT_OF_TABLECELL_UN_0)

Call LogOff()

Call FinalStatus ()










'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




