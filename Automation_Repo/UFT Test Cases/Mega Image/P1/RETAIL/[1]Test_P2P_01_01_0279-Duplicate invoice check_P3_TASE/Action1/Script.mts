
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_01_01_060-Copy Purchasing Info Records - transfer inactive article to new vendor
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
'.................Test Script Name :[1]Test_P2P_01_01_0279-Duplicate invoice check_P3
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




gstrTestCaseName = "[1]Test_P2P_01_01_0279_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0279-Duplicate invoice check_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)

Call PressEnter()

Call SetTextBox("Invoice Document No\.","RBKPV-BELNR",0,DT_MR8M_0300_INVOICE_DOCUMENT_NO,False)

Call SetTextBox("Fiscal Year","RBKPV-GJAHR",0,DT_MR8M_0300_FISCAL_YEAR,False)

Call SetTextBox("Reversal Reason","UF05A-STGRD",0,DT_MR8M_0300_REVERSAL_REASON,False)

Call TakeScreenShot()

Call PressEnter()

Call ClickButton("Reverse   \(Ctrl\+S\)",False)

Call TakeScreenShot()

Call GetStatusBar("item1","DT_MR8M_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

'Call VerifyStatusBar("Document reversed with no. "&DT_MR8M_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&": Please manually clear FI documents")

Call SetTextBox("Invoice Document No\.","RBKPV-BELNR",0,DT_MR8M_0300_INVOICE_DOCUMENT_NO_OCC1,False)

Call SetTextBox("Posting Date","G_BUDAT",0,ConvertDate(DT_MR8M_0010_INVOICE_DATE),False)

Call TakeScreenShot()

Call ClickButton("Reverse   \(Ctrl\+S\)",False)

Call TakeScreenShot()

Call GetStatusBar("item1","DT_MR8M_0300_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_MR8M_0300_OKCD)

Call PressEnter()

Call TakeScreenShot()
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_ZMDPU_MB5S_0120_COMPANY_CODE)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetCombo("RM08M-VORGANG", "Invoice")

'''Call SetTextBox("Invoice date","INVFO-BLDAT",0,ConvertDate(DT_MR8M_0010_INVOICE_DATE),False)
Call SetTextBoxNoLabel("INVFO-BLDAT",0,ConvertDate(DT_MR8M_0010_INVOICE_DATE),False)

Call PressEnter()

Call SetTextBox("Reference","INVFO-XBLNR",0,DT_MR8M_0010_REFERENCE,False)

Call PressEnter()

Call  SetTextboxNoLabel("RM08M-EBELN", 0, DT_MR8M_6211_RM08MEBELN, False)

'Call SetComboByKey("RM08M-XWARE_BNK",DT_MIRO_6211_RM08MXWARE_BNK)

Call PressEnter()

Call TakeScreenShot()

Call SelectCheckbox("INVFO-XMWST", 0, "ON", False)

Call GetTextboxValue("RM08M-DIFFERENZ", 0, "DT_MR8M_6000_CHECK_TEXT_OF_BALANCE_OUTPUT", False)

Call SetTextBox("Amount","INVFO-WRBTR",0,ConvertNegativePosetive(DT_MR8M_6000_CHECK_TEXT_OF_BALANCE_OUTPUT),False)

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Simulate Document   \(Ctrl\+Shift\+F7\)",False)

Call ClickButtonIfExist("Cancel   \(F12\)",False)

Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)

Call TakeScreenShot()

Call GetStatusBar("item1","DT_MR8M_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

'Call VerifyStatusBar("Document no. "&DT_MR8M_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''










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




