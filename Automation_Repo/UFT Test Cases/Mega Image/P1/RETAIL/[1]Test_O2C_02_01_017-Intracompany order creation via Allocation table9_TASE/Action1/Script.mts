
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_02_01_017-Intracompany order creation via Allocation table9
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


gstrTestCaseName = "Test_O2C_02_01_017-Intracompany order creation via Allocation table9"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_02_01_017-Intracompany order creation via Allocation table.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''-----------------Login-----------------------''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''''--------TransactionCode-WA01----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Allocation Table Type","AUKO-AUFAR","",DT_WA01_0100_ALLOCATION_TABLE_TYPE,False)
Call SetTextbox("Purchasing Organization","AUKO-EKORG","",DT_WA01_0100_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Purchasing Group","AUKO-EKGRP","",DT_WA01_0100_PURCHASING_GROUP,False)
Call TakeScreenShot
Call ClickButton("Create Allocation Table With User Exit Article Selection   \(Ctrl\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Directory","DY_PATH","",DT_WA01_0200_DIRECTORY,True)
Call SetTextbox("File Name","DY_FILENAME","",DT_WA01_0200_FILE_NAME,True)
Call TakeScreenShot
Call PressEnter() 
Wait(10)
Call PressEnter()
Call TakeScreenShot
wait 5
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait 5
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait 5
Call SetTableData("SAPML01ATC_0502", "Delivery Date", 1, "", "", ConvertDate(DT_WA01_0502_TABLECELL_DELIVERY_DATE_0), True)
Call SetTableData("SAPML01ATC_0502", "Percent", 1, "", "", DT_WA01_0502_TABLECELL_PERCENT_0, True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call ClickButtonIfExist("Select all   \(Shift\+F8\)",False)
Call ClickButtonIfExist("Distribution centers   \(F7\)",False)
Call VerifyStatusBar(DT_WA01_0161_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)

Call GetStatusBar("item1", "DT_WA01_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Alloc tbl created with number "&DT_WA01_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT)
Call TakeScreenShot
 
'''''''''''--------TransactionCode-WA08----------''''

Call SetTcode(DT_WA01_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("P_WORD", 2, "ON", False)
Call SetTextbox("Allocation Table","ABELN_RT-LOW","",DT_WA01_0100_CHECK_TEXT_OF_STATUSBAR_OUTPUT,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Generate Documents   \(Shift\+F5\)",False)
Call TakeScreenShot


Call GetLabelContentByRefLabel("Alloc.Tab.", 35, -32, "DT_LABEL1_OUTPUT", false)
'Call GetLabelContentByRefLabel("Alloc.Tab.", 35, -48, "DT_LABEL2_OUTPUT", false)
Call GetLabelContentByRefLabel("Wrhs order",0,-32,"DT_WA01_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)
'Call GetLabelContentByRefLabel("Wrhs order",-735,-32,"DT_WA01_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)

''''''''--------TransactionCode-WA08----------''''

Call SetTcode(DT_WA01_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_WA01_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)

Call ClickButtonIfExist("Expand Header Ctrl\+F2",false)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Supplying Site", "MEPO_TOPLINE-SUPERFIELD", "", DT_WA01_1105_CHECK_TEXT_OF_SUPPLYING_SITE, False)
Call VerifyTableCellContent(1, "Article", "SAPLMEGUITC_1211", DT_WA01_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
Call VerifyTableCellContent(2, "Article", "SAPLMEGUITC_1211", DT_WA01_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
Call TakeScreenShot
Call ClickButton("Messages   \(Shift\+F9\)",False)
Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_WA01_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)
Call TakeScreenShot

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


