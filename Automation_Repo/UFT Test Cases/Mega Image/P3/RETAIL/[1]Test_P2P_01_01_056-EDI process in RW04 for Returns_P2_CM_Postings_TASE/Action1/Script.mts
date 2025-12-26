
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_056-EDI process in RW04 for Returns_P2_CM_Postings       
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


gstrTestCaseName = "Test_P2P_01_01_056-EDI_P2_CM_Postings"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_056-EDI process in RW04 for Returns_P2_CM_Postings.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''''''----SetTransactionCode- MIRO----''''''''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code","RO02")
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetComboByKey("RM08M-VORGANG", DT_MIRO_6000_TRANSACTION)
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_DOCUMENT_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call TakeScreenShot
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call TakeScreenShot
Call SetTextboxNoLabel("RM08M-EBELN", "", DT_MIRO_6211_RM08MEBELN, False)
Call SetComboByKey("RM08M-XWARE_BNK", DT_MIRO_6211_RM08MXWARE_BNK)
Call PressEnter()
Call SelectCheckbox("INVFO-XMWST", 0, DT_MIRO_0010_CALCULATE_TAX, False)
Call TakeScreenShot
Call GetTextboxValue("RM08M-DIFFERENZ", "", "DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", false)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT),False)
Call PressEnter()
Call SelectTab("HEADER","Details",False)
Call TakeScreenShot
Call SetComboByKey("Doc. Type", DT_MIRO_0150_DOC_TYPE)
Call TakeScreenShot
Call SelectTab("HEADER","Basic Data",False)
Call TakeScreenShot
Call ClickButton("Simulate Document   \(Ctrl\+Shift\+F7\)",false)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT& " created")
Call TakeScreenShot

''''''''''----SetTransactionCode- MIR4----''''''''''

Call SetTcode(DT_MIRO_6000_OKCD)     
Call PressEnter() 

Call SetTextbox("Invoice Document No.","RBKP-BELNR","",DT_MIRO_6000_GET_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIRO_6150_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)
Call GetTextboxValue("BKPF-BELNR", "", "DT_ACCOUNTING_DOCUMENT_NUMBER_OUTPUT", false)
Call VerifyGridCellContentbyName("shell", 1, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContentbyName("shell", 2, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
'''Call VerifyGridCellContentbyName("shell", 3, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContentbyName("shell", 4, "Posting Key", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContentbyName("shell", 1, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContentbyName("shell", 2, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'''Call VerifyGridCellContentbyName("shell", 3, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContentbyName("shell", 4, "Account", "", DT_MIRO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

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



