
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_016-Purchasing  - Corrections - Cancellation - reversal process RW04 dry goods
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_016-Purchasing  - Corrections - Cancellation - reversal process RW04 dry goods
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)



gstrTestCaseName = "Test_P2P_01_01_016- process RW04 dry goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\DT_MR8M-Reverse PO Documents.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (2)
Call PressEnter()     ' - Line (3)

Call SetTextbox("Invoice Document No\.", "RBKPV-BELNR", 0, DT_MR8M_0300_INVOICE_DOCUMENT_NO, False)
Call SetTextbox("Reversal Reason","UF05A-STGRD",0,DT_MR8M_0300_REVERSAL_REASON,False)

Call PressEnter()

Call ClickButton("btn\[11\]",False)

Call GetStatusBar("text","DT_STATUSBAR_OUTPUT")
'Call VerifyStatusBar("DocumDocument reversed with no. "&DT_CREDIT_MEMO_NUMBER&" : Please manually clear FI documents")

Call SetTcode(DT_MR8M_0300_OKCD)
Call PressEnter()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("Invoice Document No\.","RBKP-BELNR",0,DT_MR8M_6150_INVOICE_DOCUMENT_NO,False)
Call TakeScreenShot()
Call PressEnter()

Call TakeScreenShot()
Call ClickButton("btn\[8\]",False)

Call TakeScreenShot()
Call VerifyGridCellContent("", 1, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''''Call VerifyGridCellContent("", 3, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
'''''Call VerifyGridCellContent("", 5, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)

Call VerifyGridCellContent("", 1, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
''''''Call VerifyGridCellContent("", 3, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 3, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
'''''Call VerifyGridCellContent("", 5, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)

Call SetTcode(DT_MR8M_0750_OKCD)
Call PressEnter()

Call SetComboByKey("GODYNPRO-ACTION",DT_MR8M_0010_GODYNPROACTION)

Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MR8M_2010_GODYNPROMAT_DOC, False)
Call PressEnter()
Call SetTableData("SAPLMIGOTV_GOITEM","OK",2,"","","ON",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call ClickButton("btn\[11\]",False)
Call TakeScreenShot()
Call GetStatusBar("text","DT_MR8M_OUTPUT")
'Call VerifyStatusBar("Article document " &DT_GR_REVERASL_NUMBER & " posted")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetComboByKey("GODYNPRO-ACTION",DT_MR8M_0010_GODYNPROACTION_OCC1)
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MR8M_2010_GODYNPROMAT_DOC_OCC1, False)
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD","Doc. info",False)
Call TakeScreenShot()

Call ClickButton("OK_FI_DOC",False)
Call TakeScreenShot()
''Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
''Call ClickButtonIfExist("Display Document   \(F2\)",True)
If SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguiedit("guicomponenttype:=31","name:=BKPF-BELNR","attachedtext:=Document Number","Index:=0").Exist(1) Then
	wait 1
Else
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
End If
Call VerifyGridCellContent("", 1, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("", 2, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
'Call VerifyGridCellContent("", 3, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL_OCC1)
'Call VerifyGridCellContent("", 4, "Posting Key", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL_OCC1)


Call VerifyGridCellContent("", 1, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)

Call VerifyGridCellContent("", 2, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)

'Call VerifyGridCellContent("", 3, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)
'
'Call VerifyGridCellContent("", 4, "Account", "", DT_MR8M_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR_OCC1)

Call GetTextboxValue("BKPF-BELNR","","DT_DOCUMENT_NUMBER_OUTPUT",False)

Call SetTcode(DT_MR8M_0100_OKCD)
Call PressEnter()

Call SetTextBox("Purchasing Document","P_EBELN",0,DT_MR8M_1000_PURCHASING_DOCUMENT,False)
Call TakeScreenShot()
Call PressEnter()

Call CLickButton("btn\[8\]",False)
Call TakeScreenshot()
Call PressEnter()
'Call VerifyifGuiLabelExists("PO " & DT_MR8M_1000_PURCHASING_DOCUMENT & " was deleted")
Call ClickButton("btn\[11\]",False)
''Call ClickButton("SPOP-VAROPTION1",False)
Call ClickButtonIfExist("SPOP-VAROPTION1",False)
Call GetStatusBar("text","DT_OUTPUT")
'Call VerifyStatusBar("Standard PO Retail " &DT_COPIED_PO_NUMBER& " changed")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MR8M_1000_OKCD)
Call PressEnter()

Call TakeScreenShot()
Call ClickButton("btn\[17\]",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_MR8M_0003_PUR_ORDER_OCC1,True)

Call PressEnter()
Call TakeScreenShot()

Call ClickButton("btn\[21\]",False)
Call TakeScreenShot()
Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_MR8M_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_4)
Call VerifyTableCellContent(4,"Status","SAPDV70ATC_NAST3",DT_MR8M_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_4)


Call ClickButton("btn\[3\]",False)

Call ClickButton("btn\[17\]",False)

'Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order")
Call SetTextBox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_MR8M_0003_PUR_ORDER_OCC2,True)

Call PressEnter()
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL", "Status", False)
Call GetTextBoxValue("MEPO1235-VALUE01",0,"DT_VAL_OUT",False)
'''Call VerifyTextBoxContent("RON","MEPO1235-VALUE01",0,DT_MR8M_1234_CHECK_TEXT_OF_ORDERED,False)
Call VerifyTextBoxContent("RON","MEPO1235-VALUE01",0,"0,00 ",False)
' VerifyTextBoxContentIconName(textboxName, textboxIndex, expectedValue, blnIsItPopup)
'Call VerifyTextBoxContentIconName

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




