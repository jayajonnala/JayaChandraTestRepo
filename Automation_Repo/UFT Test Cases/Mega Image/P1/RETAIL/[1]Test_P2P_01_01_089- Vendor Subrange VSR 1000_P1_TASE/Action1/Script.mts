
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
'.................Test Script Name :[1]Test_P2P_01_01_089- Vendor Subrange VSR 1000_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




gstrTestCaseName = "Test_P2P_01_01_089-1000_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_089- Vendor Subrange VSR 1000_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBoxNoLabel("EINA-LIFNR",0,DT_ME13_0100_VENDOR,False)

Call SetTextBox("Article","EINA-MATNR",0,DT_ME13_0100_ARTICLE,False)

Call SetTextBox("Purchasing Org\.","EINE-EKORG",0,DT_ME13_0100_PURCHASING_ORG,False)

Call PressEnter()

Call TakeScreenShot()

Call VerifyTextBoxContent("Vendor Subrange", "EINA-LTSNR", 0, DT_ME13_0101_CHECK_TEXT_OF_VENDOR_SUBRANGE, False)

Call ClickButton("Back   \(F3\)",False)

Call SetTextBox("Article","EINA-MATNR",0,DT_ME13_0100_ARTICLE_OCC1,False)

Call TakeScreenShot()

Call PressEnter()

'Call VerifyTextBoxContent("Vendor Subrange", "EINA-LTSNR", 0, DT_ME13_0101_CHECK_TEXT_OF_VENDOR_SUBRANGE_OCC1, False)

Call TakeScreenShot()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SetTcode(DT_ME13_0101_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME13_1105_MEPO_TOPLINEBSART)

Call TakeScreenShot()

Call SetTextBoxNoLabel("MEPO_TOPLINE-SUPERFIELD",0,DT_ME13_1105_VENDOR,False)
    
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME13_1221_PURCH_ORG,False) 

Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME13_1221_PURCH_GROUP,False)

Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME13_1221_COMPANY_CODE,False)     
   
Call PressEnter() 

Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME13_1211_TABLECELL_ARTICLE_0,False)    

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME13_1211_TABLECELL_PO_QUANTITY_0,False)    

Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME13_1211_TABLECELL_SITE_0,False) 

'Call SetTableData("SAPLMEGUITC_1211","Net Price","1","","",DT_ME13_1211_TABLECELL_PO_PRICE,False) 

Call PressEnter()

Call PressEnter()

Call PressEnter()

Call TakeScreenShot()

Call ClickButton("Check   \(Ctrl\+Shift\+F3\)",False)

Call TakeScreenShot()

Call VerifyStatusBar(DT_ME13_0014_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call TakeScreenShot()

Call GetStatusBar("item2","DT_ME13_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

'Call VerifyStatusBar("Standard PO Retail created under the number "&DT_ME13_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_ME13_0014_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Purchasing Document","S_EBELN-LOW",0,DT_ME13_0014_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT,False)

Call SetTextBox("Company Code","S_BUKRS-LOW",0,DT_ME13_1000_COMPANY_CODE,False)

Call SetTextBox("Partner Function","S_PARVW-LOW",0,DT_ME13_1000_PARTNER_FUNCTION,False)

Call TakeScreenShot()

Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call VerifyGridCellContent("", 1, "Vendor Subrange", 0, DT_ME13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LTSNR)

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

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




