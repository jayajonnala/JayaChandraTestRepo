
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P1
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
'.................Test Script Name :Test_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_0234-cancellation MVT 985_Fresh_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call PressEnter()
Call TakeScreenShot

Call SetTextBoxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextBoxNoLabel("GODYNPRO-DOC_YEAR",0,DT_YEAR,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Open detail data",False)

Call SelectTab("TS_GOITEM","Where", False)
Call TakeScreenShot
Call SetTextBox("Reason for Movement","GOITEM-GRUND",0,DT_MIGO_0325_REASON_FOR_MOVEMENT,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Reason for Mvmt",2,"","",DT_MIGO_0325_REASON_FOR_MOVEMENT,False)
Call SetTableData("SAPLMIGOTV_GOITEM","Reason for Mvmt",3,"","",DT_MIGO_0325_REASON_FOR_MOVEMENT,False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK",2,"","","ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK",3,"","","ON",False)
Call SelectCheckBox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call PressEnter()
Call TakeScreenShot


Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")

Call SetTcode(DT_MIGO_0001_OKCD)     
Call PressEnter() 
Call TakeScreenShot()

Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
'''''''''''Call SetTextbox("Art\. Doc\. Year","RM07M-MJAHR","",DT_MB03_0460_ART_DOC_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Accounting Documents   \(F7\)",false)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
WAIT 5
Call ClickButton("Display Document   \(F2\)",True)
Call GetTextboxValue("BKPF-BELNR",0,"DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 5, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("", 6, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("", 5, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("", 6, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)

Call VerifyGridCellContent("", 1, "XNEGP", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XNEGP)
Call VerifyGridCellContent("", 2, "XNEGP", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XNEGP)
Call VerifyGridCellContent("", 3, "XNEGP", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XNEGP)
Call VerifyGridCellContent("", 4, "XNEGP", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_XNEGP)
Call VerifyGridCellContent("", 5, "XNEGP", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_XNEGP)
Call VerifyGridCellContent("", 6, "XNEGP", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_XNEGP)

Call LogOff()

Call FinalStatus()






'''''
'''''Call SetTextBox("Company Code","S_BUKRS-LOW",0,DT_ZMDIM_DESTR_STOCK_1000_COMPANY_CODE,False)
'''''Call SetTextBox("Site","S_WERKS-LOW",0,DT_ZMDIM_DESTR_STOCK_1000_SITE,False)
'''''Call SetTextBox("Storage Location","P_LGORT",0,DT_ZMDIM_DESTR_STOCK_1000_STORAGE_LOCATION,False)
'''''Call SetTextBox("Movement type","P_BWART",0,DT_ZMDIM_DESTR_STOCK_1000_MOVEMENT_TYPE,False)
'''''Call SetTextBox("Customer","P_KUNNR",0,DT_ZMDIM_DESTR_STOCK_1000_CUSTOMER,False)
'''''
'''''
'''''Call SetTextBox("Location","P_ABLAD",0,DT_ZMDIM_DESTR_STOCK_1000_LOCATION,False)
'''''
'''''Call SetTextBox("Reason for Movement","P_GRUND",0,DT_ZMDIM_DESTR_STOCK_1000_REASON_FOR_MOVEMENT,False)
'''''
'''''
'''''Call TakeScreenShot()
'''''
'''''Call ClickButton("%_S_MATNR_%_APP_%-VALU_PUSH",False)
'''''
'''''
'''''Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_0, False)
'''''
'''''Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_1, False)
'''''
'''''Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_2, False)
'''''
'''''Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_3, False)
'''''
'''''Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_4, False)
'''''
'''''Call SetTableData("SAPLALDBSINGLE", "Single value", 6, "", "", DT_ZMDIM_DESTR_STOCK_3010_TABLECELL_SINGLE_VALUE_5, False)
'''''
'''''Call ClickButton("Copy   \(F8\)",False)
'''''
'''''Call ClickButton("Execute   \(F8\)",False)
'''''
'''''
'''''Call SetGridData("",1, "Quantity",DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_0_QUANTITY, False)
'''''
'''''Call SetGridData("",2, "Quantity",DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_1_QUANTITY, False)
'''''
'''''Call SetGridData("",3, "Quantity",DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_2_QUANTITY, False)
'''''
'''''Call SetGridData("",4, "Quantity",DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_3_QUANTITY, False)
'''''
'''''Call SetGridData("",5, "Quantity",DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_4_QUANTITY, False)
'''''
'''''Call SetGridData("",6, "Quantity",DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_5_QUANTITY, False)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call CLickButton("Select all   \(F5\)",False)
'''''
'''''Call CLickButton("btn\[8\]",False)
'''''
'''''Call TakeScreenShot()
'''''
'''''
'''''Call GetGridContent("", 0, "Created Document", 1, "", "", DT_ZMDIM_DESTR_STOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OUTPUT)
'''''
'''''Call SetTcode(DT_ZMDIM_DESTR_STOCK_0500_OKCD)
'''''
'''''Call PressEnter()
'''''
''''''Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_ZMDIM_DESTR_STOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OUTPUT,False)
'''''
''''''Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_ZMDIM_DESTR_STOCK_1000_ARTICLE_DOC_YEAR,False)
'''''
''''''Call SetTextBox("Sort order","PM_NSORT",0,DT_ZMDIM_DESTR_STOCK_1000_SORT_ORDER,False)
'''''
''''''Call SetTextBox("Processing Mode","PM_VERMO",0,DT_ZMDIM_DESTR_STOCK_1000_PROCESSING_MODE,False)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call ClickButton("Execute   \(F8\)",False)
''''''' SelectCheckboxNoLabel(checkBoxIndex, OnOffStatus, blnIsItPopup)
''''''SelectCheckboxNoLabel
''''''
'''''
'''''Call SelectCheckBoxByGuiLabel(DT_ZMDIM_DESTR_STOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OUTPUT, -17, -1, "ON")
'''''
'''''Call ClickButton("Print preview   \(Shift\+F4\)",False)
'''''
'''''Call TakeScreenShot()
'''''
'''''Call SetTcode(DT_ZMDIM_DESTR_STOCK_0100_OKCD)
'''''
'''''Call PressEnter()
'''''
'''''Call SetTextBox("Company Code","BUKRS-LOW",0,DT_ZMDIM_DESTR_STOCK_1000_COMPANY_CODE_OCC1,False)
'''''
'''''Call SetTextBox("Article","MATNR-LOW",0,DT_ZMDIM_DESTR_STOCK_1000_ARTICLE_DOCUMENT_OCC1,False)
'''''
'''''Call PressEnter()
'''''
'''''Call ClickButton("Execute   \(F8\)",False)
'''''
'''''Call TakeScreenShot()


'Call VerifyGridCellContent("", 1,"Movement type", "", DT_ZMDIM_DESTR_STOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)

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




