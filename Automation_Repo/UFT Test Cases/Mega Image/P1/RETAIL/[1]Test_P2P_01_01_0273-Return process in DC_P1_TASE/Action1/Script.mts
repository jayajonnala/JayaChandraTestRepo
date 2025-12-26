
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
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_P2P_01_01_0273_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\DLL\DT_P2P_01_01_0273-Return process in DC_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()

Call  SetComboByKey("MEPO_TOPLINE-BSART", DT_ME21N_1105_MEPO_TOPLINEBSART)

Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD",0,DT_ME21N_1105_VENDOR,False)  

Call PressEnter()

Call SetTextbox("Purch. Org.","MEPO1222-EKORG",0,DT_ME21N_1221_PURCH_ORG,False)     ' 

Call SetTextbox("Purch. Group","MEPO1222-EKGRP",0,DT_ME21N_1221_PURCH_GROUP,False)     '

Call SetTextbox("Company Code","MEPO1222-BUKRS",0,DT_ME21N_1221_COMPANY_CODE,False)     
   
Call PressEnter()    

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)

Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)

Call SetTableData("SAPLMEGUITC_1211","Stor. Location","1","","",DT_ME21N_1211_TABLECELL_STOR_LOCATION_0,False)


Call SetTableData("SAPLMEGUITC_1211", "Returns Item", 1, "", "", "ON", False)


Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)

Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)

Call SetTableData("SAPLMEGUITC_1211","Stor. Location","2","","",DT_ME21N_1211_TABLECELL_STOR_LOCATION_1,False)
Call SetTableData("SAPLMEGUITC_1211", "Returns Item", 2, "", "", "ON", False)

Call TakeScreenShot()

Call PressEnter()

Call PressEnter()

Call ClickButton("Check   \(Ctrl\+Shift\+F3\)",False)

Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetStatusBar("item2","DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")

Call VerifyStatusBar("Standard PO Retail created under the number "&DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_ME21N_0014_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call CLickButton("Other Purchase Order   \(Shift\+F5\)",False)

Call SetTextBox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT,False)

Call ClickButton("Other Document   \(Enter\)",False)

Call TakeScreenShot()

Call ClickButton("Messages   \(Shift\+F9\)",False)

Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_ME21N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1, "Output Type", "SAPDV70ATC_NAST3", DT_ME21N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_ME21N_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Document Number","S_EBELN-LOW",0,DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT,False)

Call SetTextBox("Application","P_KAPPL",0,DT_ME21N_1000_APPLICATION,False)

Call SetTextBox("Purchasing Organization","S_EKORG-LOW",0,DT_ME21N_1000_PURCHASING_ORGANIZATION,False)

Call SetTextBox("Processing Status","P_VSTAT",0,DT_ME21N_1000_PROCESSING_STATUS,False)

Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)

Call VerifyGridCellContent("", 1, "Message type", 0, DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)

Call VerifyGridCellContent("", 2, "Message type", 0, DT_ME21N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KSCHL)
Call SelectRowGuiGrid("", 0, "Purchasing Document", DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT, False)
Call ClickButton("Display Message   \(Shift\+F8\)",False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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




