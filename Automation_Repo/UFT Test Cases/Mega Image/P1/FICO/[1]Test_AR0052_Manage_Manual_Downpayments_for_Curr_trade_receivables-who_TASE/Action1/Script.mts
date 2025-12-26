
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_AR0052_Manage_Manual_Downpayments_for_Curr_trade_receivables-who_TASE
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


gstrTestCaseName = "Test_AR0052_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''----------------------Tcode F-43----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace((DT_F43_100_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Posting Date","BKPF-BUDAT","",Replace((DT_F43_100_POSTING_DATE),"/","."),False)
Call SetTextbox("Type","BKPF-BLART","",DT_F43_100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F43_100_COMPANY_CODE,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F43_100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F43_100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F43_100_DOCHEADER_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F43_100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F43_100_ACCOUNT,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-43").SapGuiWindow("transaction:=F-43").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call GetTextboxValue("BSEG-HKONT","","DT_DR_GL_ACC_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectCheckbox("BKPF-XMWST","1",DT_F43_301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F43_301_AMOUNT,False)
''text box label changed
'''Call SetTextbox("Payt Terms","BSEG-ZTERM","",DT_F43_301_PAYT_TERMS,False)
Call SetTextboxNoLabel("BSEG-ZTERM","",DT_F43_301_PAYT_TERMS,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F43_301_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F43_301_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F43_301_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F43_301_ACCOUNT,False)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F43_301_SGL_IND,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-43").SapGuiWindow("transaction:=F-43").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

Call GetTextboxValue("BSEG-HKONT","","DT_CR_GL_ACC_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectCheckbox("RF05A-XMWST","1",DT_F43_301_CALCULATE_TAX,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_F43_304_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_F43_304_TAX_CODE,False)
Call SetTextbox("Bus\. Area","BSEG-GSBER","",DT_F43_304_BUS_AREA,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F43_304_TEXT,False)
Call SetTextbox("Profit Ctr","BSEG-PRCTR","",DT_PROFIT_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
For Iterator = 1 To 6 Step 1
	If SAPGuiSession("transaction:=F-43").SapGuiWindow("transaction:=F-43").SAPGuiStatusBar("messagetype:=W").Exist = True Then
		Call PressEnter() 
	End If
Next

'veryfy sattus bar content
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_F43_100_CHECK_TEXT_OF_STATUSBAR)
Call VerifyStatusBarMessageType("S")

'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_F43_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_F43_100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_F43_100_COMPANY_CODE_OCC2,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_F43_100_FISCAL_YEAR,False)

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display Document Header   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Document type","BKPF-BLART","",DT_F43_1710_CHECK_TEXT_OF_DOCUMENT_TYPE,False)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Call DoubleClickGuiGridCell("",0, 1, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Acc","BSEG-HKONT",0,DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC2, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",1,"HKONT","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT_OCC2)
Call DoubleClickGuiGridCell("",0, 2, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L acct","BSEG-HKONT",0,DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",2,"HKONT","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call DoubleClickGuiGridCell("",0, 3, "Company code", False)
Call TakeScreenShot
Call VerifyTextBoxContent("G/L Account","BSEG-HKONT",0,DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT, False)
Call ClickButton("Back   \(F3\)", False)
'Call VerifyGridCellContent("",3,"HKONT","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)

Call VerifyGridCellContent("",1,"BSCHL","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"BSCHL","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"BSCHL","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)

Call VerifyGridCellContent("",2,"UMSKZ","",DT_F43_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)

Call VerifyTextBoxContent("Document Number","BKPF-BELNR",0,DT_F43_750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)


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


