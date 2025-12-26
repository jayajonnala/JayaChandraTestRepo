
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCIM03_007_Pallet_File_Quantity_Surplus_in_status_10_P03 
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

gstrTestCaseName = "Test_06DCIM03_007_Surplus_10_P03_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\AB\RETAIL\DT_06DCIM03_007_Pallet_File_Quantity_Surplus_in_status_10_P03_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''''-----Login----------'''''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Site","WERKS-LOW","",DT_MB51_1000_SITE,False)
Call SetTextbox("Movement type","BWART-LOW","",DT_MB51_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Article","MATNR-LOW","",DT_MB51_1000_ARTICLE,False)
Call SetTextboxNoLabel("BKTXT-LOW","",DT_MB51_1000_DOCUMENT_HEADER_TEXT,False)
Call SetTextbox("Document Date","BLDAT-LOW","",ConvertDate(DT_MB51_1000_DOCUMENT_DATE),False)

Call TakeScreenShot
Call ClickButtonIfExist("%_LGORT_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "", "", DT_SLOC_1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_SLOC_2, True)
Call ClickButtonIfExist("Copy   \(F8\)",False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call SelectMenuBar("Settings;Layout;Choose...")
Call ClickButtonIfExist("Continue   \(Enter\)",False)
Call SelectRowGuiGridbyRowNo("", 0, 6, False)
Call SelectColumnGuiGrid("", 0, "Article Document", false)
Call ClickButtonIfExist("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)
Call TakeScreenShot

Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND", 0)

Call SetTextBox("Search Term:","GS_SEARCH-VALUE",0,"Document Header Text",True)
Call SelectCheckBox("GS_SEARCH-EXACT_WORD",0,"ON",True)
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenSHot()
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenSHot()

Call ClickButtonToolBar("&FIND", 0)
Call SetTextBox("Search Term:","GS_SEARCH-VALUE",0,"Receiving stor. loc.",True)
Call SelectCheckBox("GS_SEARCH-EXACT_WORD",0,"ON",True)
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenSHot()
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenSHot()
Call ClickButton("Transfer   \(Enter\)",True)

Call TakeScreenShot()

Call VerifyGridCellContent("", 2, "Qty in unit of entry", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ERFMG)
Call VerifyGridCellContent("", 1, "Storage Location", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LGORT)
Call VerifyGridCellContent("", 1, "Receiving stor. loc.", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMLGO)
Call VerifyGridCellContent("", 1, "Document Header Text", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BKTXT)



'''''--------TransactionCode-MMBE----------''''


Call SetTcode(DT_MB51_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Article","MS_MATNR-LOW","",DT_MB51_1000_ARTICLE_OCC1,False)
Call SetTextbox("Site","MS_WERKS-LOW","",DT_MB51_1000_SITE_OCC1,False)
Call TakeScreenShot
Call ClickButton("%_MS_LGORT_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 1, "<NA>", "<NA>", DT_SLOC_1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "<NA>", "<NA>", DT_SLOC_2, True)
Call ClickButton("btn\[8\]",True)
Call ClickButton("Execute   \(F8\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#1")
''Call GetGridContent("SLoc", 0, "Stock", 1, "<NA>", "<NA>", "DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE_OCC1_OUTPUT")
'Call VerifyGridCellContent("SLoc", 1, "Stock", 0, DT_MB51_0700_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSTNDTXT)
Call ClickButton("Continue   \(Enter\)",False)
Call ActivateNodeGuiTree(0, "#1;#1;#1;#2")
''Call GetGridContent("SLoc", 0, "Stock", 1, "<NA>", "<NA>", "DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE_OCC3_OUTPUT")
'Call VerifyGridCellContent("SLoc", 1, "Stock", 0, DT_MB51_0700_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSTNDTXT_OCC1)
Call ClickButton("Continue   \(Enter\)",False)
Call TakeScreenShot
Call VerifyNodeTextGuiTree(0, "0001 Εμπορευσίμων")
Call VerifyNodeTextGuiTree(0, (DT_MB51_0300_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC1))

Call LogOff()
Call FinalStatus()






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




