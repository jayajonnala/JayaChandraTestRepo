
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_024-Import from EU Vendors GR before Invoice_P6_Inv_del_cost     
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

gstrTestCaseName = "Test_P2P_01_01_024- Invoice_P6_Inv_del_cost"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_024-Import from EU Vendors GR before Invoice_P6_Inv_de.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetComboByKey("RM08M-VORGANG", DT_MIRO_6000_TRANSACTION)
Call PressEnter()
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)

Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
 Call TakeScreenShot
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call PressEnter() 
Call TakeScreenShot
Call SetComboByKey("RM08M-XWARE_BNK", DT_MIRO_6211_RM08MXWARE_BNK)
Call PressEnter() 
Call SetTextBoxNoLabel("RM08M-EBELN",0,DT_MIRO_6211_RM08MEBELN,False)
Call PressEnter()
Call SetTableData("SAPLMR1MTC_MR1M","Tax Code",1,"","",DT_MIRO_6310_TABLECELL_TAX_CODE_0,False)
Call SetTableData("SAPLMR1MTC_MR1M","Tax Code",2,"","",DT_MIRO_6310_TABLECELL_TAX_CODE_1,False)
Call SelectCheckBox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call GetStatusBar("item1","DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
Call SelectMenuBar("Invoice Document;Display")
Call TakeScreenShot

Call VerifyTextBoxNoLabelContent("RBKPV-BELNR",0,DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call TakeScreenShot
Call GetTextboxValue("BKPF-BELNR",0,"DT_MIRO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call SelectMenuIdToolBar("&COL0",0)
''Call SelectCellGuiGrid("Column Set",0,129,"Column Name",True)
Call ClickButtonToolBar("&FIND",True)
Call SetTextboxNoLabel("GS_SEARCH-VALUE","",DT_MIRO_0841_SEARCH_TERM,True)
Call SelectCheckBox("GS_SEARCH-EXACT_WORD",0,"ON",True)
Call PressEnter() 
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 5, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)

Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 5, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

Call VerifyGridCellContent("", 1, "KTOSL", 0,DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTOSL)
Call VerifyGridCellContent("", 2, "KTOSL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTOSL)
Call VerifyGridCellContent("", 3, "KTOSL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTOSL)
Call VerifyGridCellContent("", 5, "KTOSL", 0,DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTOSL)


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


