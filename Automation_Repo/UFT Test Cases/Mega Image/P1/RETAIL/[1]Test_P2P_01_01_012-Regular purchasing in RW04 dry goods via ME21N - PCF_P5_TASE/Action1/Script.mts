
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_012-Regular purchasing in RW04 dry goods via ME21N - PCF_P5
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

gstrTestCaseName = "Test_P2P_01_01_012- RW04 dry goods via ME21N - PCF_P5"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_012-Regular purchasing in RW04  dry goods  via ME21N - PCF_P5.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''--------TransactionCode-MIRO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetCombo("RM08M-VORGANG", DT_INVOICE)
Call ClickButton("Back   \(F3\)",False)
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call PressEnter()  
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter() 
Call SetCombo("RM08M-REFERENZBELEGTYP","Purchase Order/Scheduling Agreement")
Call TakeScreenShot
Call SetTextboxNoLabel("RM08M-EBELN", 0, DT_MIRO_6211_RM08MEBELN, False)
Call PressEnter()

Call SelectCheckbox("INVFO-XMWST", 0, DT_MIRO_0010_CALCULATE_TAX, False)
Call GetTextboxValue("RM08M-DIFFERENZ", 0, "DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",DT_MIRO_6000_CHECK_TEXT_OF_BALANCE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("btn\[43\]", False)
Call VerifyTableCellContentPopup(1, "G/L", "SAPLMR1MTC_MR1M_PB", DT_MIRO_6250_CHECK_TEXT_OF_TABLECELL_GL_0)
Call VerifyTableCellContentPopup(2, "G/L", "SAPLMR1MTC_MR1M_PB", DT_MIRO_6250_CHECK_TEXT_OF_TABLECELL_GL_1)
Call VerifyTableCellContentPopup(3, "G/L", "SAPLMR1MTC_MR1M_PB", DT_MIRO_6250_CHECK_TEXT_OF_TABLECELL_GL_2)
Call TakeScreenShot

Call ClickButton("btn\[11\]", False)

Call GetStatusBar("item1", "DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT" )
Call VerifyStatusBar("Document no. "&DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot

Call SelectMenuBar("Invoice Document;Display")
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)", False)

Call GetTextboxValue("BKPF-BELNR", 0, "DT_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT",DT_ACCOUNTING_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

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




