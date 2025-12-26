
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_AR0058 Manage Manual Guarantees for Other receivables - rent_TASE
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


gstrTestCaseName = "AR0058"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AR0001 Manage AR Documents-Manage Manual Customer Invoicing_Cred.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''----------------------Tcode FD03----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Customer","RF02D-KUNNR","",DT_FD03_7106_CUSTOMER,True)
Call SetTextbox("Company Code","RF02D-BUKRS","",DT_FD03_7106_COMPANY_CODE,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Customer","RF02D-KUNNR","",DT_FD03_7002_CHECK_TEXT_OF_CUSTOMER,False)
Call ClickButton("Company Code Data   \(Ctrl\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextboxValue("KNB1-AKONT","","DT_FD03_7211_CHECK_TEXT_OF_RECON_ACCOUNT_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Recon\. account","KNB1-AKONT","",DT_FD03_7211_CHECK_TEXT_OF_RECON_ACCOUNT,False)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Cancel   \(F12\)",True)

'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
'''----------------------Tcode F-22----------------------------
'Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_FD03_0100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_FD03_0100_POSTING_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_FD03_0100_TYPE,False)
Call SetTextbox("Period","BKPF-MONAT","",DT_FD03_0100_PERIOD,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FD03_0100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FD03_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FD03_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_FD03_0100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FD03_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FD03_0100_ACCOUNT,False)
'Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F22_0100_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SelectCheckbox("BKPF-XMWST","1",DT_FD03_0301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FD03_0301_AMOUNT,False)
'Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_F22_0301_PAYT_TERMS,False)
'Call SetTextbox("Due on","BSEG-ZFBDT","",Replace((DT_F22_0303_DUE_ON),"/","."),False)
'Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F22_301_BUS_AREA,False)
'Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0301_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FD03_0301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_FD03_0301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_FD03_0301_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_FD03_0301_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call SetTextbox("Amount","BSEG-WRBTR","",DT_FD03_0304_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_FD03_0304_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_FD03_0304_TAX_CODE,False)
'Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F22_0300_ASSIGNMENT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_FD03_0304_TEXT,False)
Call SetTextbox("Profit Ctr","BSEG-PRCTR","",DT_PROFIT_CENTER,False)
'Call SetTextbox("Business Area","COBL-GSBER","",DT_F22_1006_BUSINESS_AREA,False)
'Call SetTextbox("Cost Center","COBL-KOSTL","",DT_F22_1006_COST_CENTER,False)
'Call SetTextbox("Profit Center","COBL-PRCTR","",DT_F22_1006_PROFIT_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-22").SapGuiWindow("transaction:=F-22").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_FD03_0100_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Yes",True)
'Capture the screenshot
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_FD03_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FD03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FD03_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FD03_0100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call GetGridContentByTitle("","","Assignment",1,"DT_FD03_0750_CHECK_TEXT_OF_REFERENCE_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call DoubleClickGuiGridCell("",0, 2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L acct","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",2,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call DoubleClickGuiGridCell("",0, 3, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",3,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)

Call VerifyGridCellContent("",1,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"BSCHL","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

'Call VerifyGridCellContent("",2,"UMSKZ","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)

Call VerifyGridCellContent("",1,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("",1,"LOKKT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("",2,"LOKKT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("",3,"LOKKT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LOKKT)

Call VerifyGridCellContent("",1,"ZUONR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"ZUONR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)

Call VerifyGridCellContent("",1,"SGTXT","",DT_FD03_0301_CHECK_TEXT_OF_TEXT)
Call VerifyGridCellContent("",2,"SGTXT","",DT_FD03_0304_CHECK_TEXT_OF_TEXT)

Call VerifyGridCellContent("",1,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",3,"AZBET","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

Call VerifyTextBoxContent("Currency","BKPF-WAERS","",DT_FD03_0750_CHECK_TEXT_OF_CURRENCY,False)

'Call ClickButton("Document Display: General Ledger View   \(Ctrl\+F9\)",False)
''Capture the screenshot
'Call TakeScreenShot()

'Call VerifyGridCellContent("",1,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC1)
'Call VerifyGridCellContent("",2,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT_OCC1)
'Call VerifyGridCellContent("",3,"HKONT","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT_OCC1)

'Call VerifyGridCellContent("",1,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
'Call VerifyGridCellContent("",2,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
'Call VerifyGridCellContent("",3,"Account","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR_OCC1)

'Call VerifyGridCellContent("",1,"ZUONR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR_OCC1)
'Call VerifyGridCellContent("",2,"ZUONR","",DT_FD03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR_OCC1)

Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_FD03_0750_CHECK_TEXT_OF_REFERENCE,False)


Wait(1)
'Log Off From Applicaton
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




