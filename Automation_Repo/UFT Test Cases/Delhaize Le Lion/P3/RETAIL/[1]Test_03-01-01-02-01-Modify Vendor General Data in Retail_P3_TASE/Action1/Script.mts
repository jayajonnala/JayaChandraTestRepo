
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : 
'.................Test Scenario: Test_03-01-01-02-01-Modify Vendor General Data in Retail_P3
'.................TCode: XK01
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_03-01-01-02-01-GD_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'ensure no open session
Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-XK02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextboxNoLabel("RF02K-LIFNR","",DT_XK02_0101_VENDOR,False)
'This function SetTextboxNoLabel is used for vendor to supplier change.
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK02_0101_COMPANY_CODE,False)
Call SetTextbox("Purchasing Organization","RF02K-EKORG","",DT_XK02_0101_PURCH_ORGANIZATION,False)
Call TakeScreenShot
Call ClickButton("Select All   \(F7\)",False)
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call SelectCheckbox("LFM1-BOLRE",1,DT_XK02_0310_SUBSEQUENT_SETTLEMENT, False)
Call SelectCheckbox("LFM1-BOIND",1,DT_XK02_0310_SUBSEQ_SETT_INDEX, False)
Call SelectCheckbox("LFM1-KZAUT",1,DT_XK02_0310_AUTOMATIC_PURCHASE_ORDER, False)
Call SelectCheckbox("LFM1-XNBWY",1,DT_XK02_0310_REVALUATION_ALLOWED, False)
Call SelectCheckbox("LFM1-NRGEW",1,DT_XK02_0310_GRANT_DISCOUNT_IN_KIND, False)
Call SelectCheckbox("LFM1-PRFRE",1,DT_XK02_0310_RELEVANT_FOR_PRICE_DET__VENDOR_HIERARCHY, False)
Call SelectCheckbox("LFM1-AGREL",1,DT_XK02_0310_RELEVANT_FOR_AGENCY_BUSINESS, False)
Call SetTextbox("Incoterms","LFM1-INCO1","",DT_XK02_0310_INCOTERMS,False)
Call SetTextbox("Incoterms","LFM1-INCO2","",DT_XK02_0310_INCOTERMS_OCC1,False)
Call SetTextbox("Pricing Date Control","LFM1-MEPRF","",DT_XK02_0310_PRICING_DATE_CONTROL,False)
Call SetTextbox("PROACT control prof.","LFM1-PAPRF","",DT_XK02_0310_PROACT_CONTROL_PROF,False)
Call SetTextbox("Purchasing group","LFM1-EKGRP","",DT_XK02_0310_PURCHASING_GROUP,False)
Call SetTextbox("Unit of measure grp","LFM1-MEGRU","",DT_XK02_0310_UNIT_OF_MEASURE_GRP,False)
Call TakeScreenShot
Call ClickButton("Sub-ranges   \(Shift\+F4\)",False)
Call SetTableData("SAPMF02KTCTRL_TEILSORTIMENTE", "Vendor Subrange",1,"", "",DT_XK02_0360_TABLECELL_VENDOR_SUBRANGE_0, False)
Call SetTableData("SAPMF02KTCTRL_TEILSORTIMENTE", "VSR description",1,"", "",DT_XK02_0360_TABLECELL_VSR_DESCRIPTION_0, False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Different Data   \(Shift\+F6\)", False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call SelectCheckbox("WRF02K-HINZE",1,DT_XK02_0352_PURCHASING_DATA, True)
Call SelectCheckbox("WRF02K-HINZP",1,DT_XK02_0352_PARTNER_FUNCTIONS, True)
Call SetTextbox("Vendor Subrange","WRF02K-LTSNR","",DT_XK02_0352_VENDOR_SUBRANGE,True)
Call SetTextbox("VSR description","WYT1T-LTSBZ","",DT_XK02_0352_VSR_DESCRIPTION,True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButton("Back   \(F3\)", False)
Call ClickButtonIfExist("Partner details   \(F7\)",False)
Call TakeScreenShot
Call SetTableData("SAPMF02KTCTRL_PARTNERROLLEN", "Number",1,"", "",DT_XK02_0320_TABLECELL_NUMBER_0, False)
Call SetTableData("SAPMF02KTCTRL_PARTNERROLLEN", "Number",4,"", "",DT_XK02_0320_TABLECELL_NUMBER_3, False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar(DT_XK02_0101_CHECK_TEXT_OF_STATUSBAR)
Call LogOff()
Call FinalStatus ()
'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

