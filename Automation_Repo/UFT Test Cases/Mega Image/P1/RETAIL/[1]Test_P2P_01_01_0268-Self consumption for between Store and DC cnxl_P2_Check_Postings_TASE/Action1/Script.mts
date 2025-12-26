
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
'.................Test Script Name :[1]Test_P2P_01_01_0268-Self consumption for between Store and DC cnxl_P1_SO_OD_GI_INV
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



'The test case has been modified due to execution issue.

gstrTestCaseName = "Test_P2P_01_01_0268_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0268-Self consumption for between Store and DC_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (2)
Call PressEnter()     ' - Line (3)

Call SetTextBox("Order","VBAK-VBELN",0,DT_VA03_0102_ORDER,False)

Call TakeScreenShot()

Call PressEnter()

Call VerifyTableCellContent(1, "Article", "SAPMV45ATCTRL_U_ERF_AUFTRAG", DT_VA03_4900_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)

Call TakeScreenShot()

Call CLickButton("BT_PKON",False)

Call SelectTab("TAXI_TABSTRIP_ITEM","Shipping",False)

Call VerifyTextBoxContent("Storage Loc\.", "VBAP-LGORT", 0, DT_VA03_4452_CHECK_TEXT_OF_STOR_LOCATION, False)

Call TakeScreenShot()

Call ClickButton("btn\[3\]",False)

Call ClickButton("Display document flow   \(F5\)",False)

Call TakeScreenShot()

Call ActivateNodeGuiTree(0, "#1;#1")

Call DoubleClick()

Call VerifyNodeTextGuiTree(0,DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1)


Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call DoubleClick()

Call GetGridContent("Self.*", 0, "Doc.no.", 1,"Number of Preceding Document", DT_IMPORT_OUTBOUND_DELIVERY_NUMBER, "DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC1_OUTPUT")



Call ActivateNodeGuiTree(0,"#1;#1;#2")

Call DoubleClick()

Call VerifyNodeTextGuiTree(0,DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC2)
Call ActivateNodeGuiTree(0,"#1;#1;#2;#1")

Call DoubleClick()
Call GetGridContent("Accounting.*", 0, "Doc.no.", 1, "Number of Preceding Document", DT_IMPORT_OUTBOUND_DELIVERY_NUMBER, "DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC3_OUTPUT")

Call TakeScreenShot()

Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call ClickButton("Display document   \(F8\)",False)



Call GetTextboxValue("RM07M-MTSNR", 0, "DT_VA03_0420_CHECK_TEXT_OF_ART_SLIP_OUTPUT", False)

Call TakeScreenShot()

Call VerifyTextBoxContent("Posting Date", "MKPF-BUDAT", 0,ConvertDate( DT_VA03_0420_CHECK_TEXT_OF_POSTING_DATE), False)

Call ClickButton("Accounting Documents   \(F7\)",False)

Call TakeScreenShot()

Call GetTextboxValue("BKPF-BELNR",0,"DT_VA03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)


Call  VerifyGridCellContent("", 1, "Posting Key", 0, DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call  VerifyGridCellContent("", 2, "Posting Key", 0, DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call  VerifyGridCellContent("", 1, "Account", 0, DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call  VerifyGridCellContent("", 2, "Account", 0, DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call  VerifyGridCellContent("", 1, "Profit Center", 0, DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call  VerifyGridCellContent("", 2, "Profit Center", 0, DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

Call ClickButton("Back   \(F3\)",False)

Call FocusTextBox("Item", "MSEG-ZEILE", False)

Call ClickButton("Choose   \(F2\)",False)

Call CLickButton("Output   \(Shift\+F2\)",False)

Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_VA03_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SetTcode(DT_VA03_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()



Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC1_OUTPUT,False)

Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_YEAR,False)

Call SetTextBox("Sort order","PM_NSORT",0,DT_VA03_1000_SORT_ORDER,False)

Call SetTextBox("Processing Mode","PM_VERMO",0,DT_VA03_1000_PROCESSING_MODE,False)

Call TakeScreenShot()

Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call SelectCheckBoxNoLabel(0,"ON",False)


Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call TakeScreenShot()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_VA03_0100_OKCD_OCC1)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Billing document","VBRK-VBELN",0,DT_VA03_0101_BILLING_DOCUMENT,False)

Call PressEnter()

Call CLickButton("Accounting overview   \(Shift\+F4\)",False)

Call TakeScreenShot()

Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", False)

Call ClickButton("Display Document   \(F2\)",False)

Call TakeScreenSHot()

Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call VerifyGridCellContent("", 2, "Account", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 3, "Account", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("", 1, "Profit Center", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call VerifyGridCellContent("", 2, "Profit Center", 0, DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)


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




