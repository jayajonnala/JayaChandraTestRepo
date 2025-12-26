
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_P2P_01_01_089- VSR 1000_P3_TASE"
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
'Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextBoxNoLabel("LIFNR-LOW",0,DT_ZMDPU_MB5S_1000_VENDOR,False)

Call SetTextBox("Purch\. Organization","EKORG-LOW",0,DT_ZMDPU_MB5S_1000_PURCH_ORGANIZATION,False)

Call SetTextBox("Purchasing Document","EBELN-LOW",0,DT_ZMDPU_MB5S_1000_PURCHASING_DOCUMENT,False)

Call SetTextBox("Partner Function","S_PARVW-LOW",0,DT_ZMDPU_MB5S_1000_PARTNER_FUNCTION,False)

Call SetTextBox("Company Code","S_BUKRS-LOW",0,DT_ZMDPU_MB5S_0120_COMPANY_CODE,False)

Call TakeScreenShot()

Call PressEnter()

'Call ClickButton("Continue   \(Enter\)",True)

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call SetTcode(DT_ZMDPU_MB5S_0120_OKCD)     

Call PressEnter() 

Call TakeScreenShot()
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_ZMDPU_MB5S_0120_COMPANY_CODE)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetComboByKey("RM08M-VORGANG", DT_ZMDPU_MB5S_6000_TRANSACTION)
Call PressEnter()
Call SetTextBoxNoLabel("INVFO-BLDAT",0,ConvertDate(DT_ZMDPU_MB5S_0010_INVOICE_DATE),False)
Call PressEnter()
Call SetTextBox("Reference","INVFO-XBLNR",0,DT_ZMDPU_MB5S_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_ZMDPU_MB5S_6000_TRANSACTION)
Call PressEnter() 
Call SetComboByKey("RM08M-XWARE_BNK", DT_ZMDPU_MB5S_6000_TRANSACTION)
Call PressEnter() 
Call TakeScreenShot
Call SetTextboxNoLabel("RM08M-EBELN", 0, DT_ZMDPU_MB5S_6211_RM08MEBELN, False)

Call PressEnter()

Call TakeScreenShot()

Call SelectTab("HEADER", "Payment", False)

Call TakeScreenShot()
'this step is no more used by business Defect #20238
'' Call ClickButton("Other bank details",False)
'' 
'' Call TakeScreenShot()
 
'' Call ClickButton("Back   \(F3\)",False)
 
 Call SelectTab("HEADER", "Basic Data", False)

Call SelectCheckbox("INVFO-XMWST", 0, "ON", False)

Call GetTextboxValue("RM08M-DIFFERENZ", 0, "DT_ZMDPU_MB5S_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", False)

Call SetTextBox("Amount","INVFO-WRBTR",0,ConvertNegativePosetive(DT_ZMDPU_MB5S_6000_CHECK_TEXT_OF_BALANCE_OUTPUT),False)

Call PressEnter()

Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_ZMDPU_MB5S_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

Call SelectMenuBar("Invoice Document;Display")

'Call ClickButton("Other bank details",False)

Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False)

Call  GetTextboxValue("BKPF-BELNR", 0, "DT_ZMDPU_MB5S_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", False)

Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_ZMDPU_MB5S_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_ZMDPU_MB5S_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_ZMDPU_MB5S_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_ZMDPU_MB5S_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call VerifyGridCellContent("", 2, "Account", 0, DT_ZMDPU_MB5S_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 3, "Account", 0, DT_ZMDPU_MB5S_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call LogOff()

Call FinalStatus()
















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




