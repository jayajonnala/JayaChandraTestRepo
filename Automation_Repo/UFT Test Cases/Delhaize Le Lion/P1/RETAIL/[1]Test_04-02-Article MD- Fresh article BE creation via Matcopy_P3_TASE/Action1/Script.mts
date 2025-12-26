

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04-02-Article MD- Fresh article BE creation via Matcopy_P3
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

gstrTestCaseName = "Test_04-02-Article MD- Fresh article BE creation via Matcopy_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MM42 -----------------------------------



Call SetTextbox("Article","RMMW1-MATNR","",DT_MM42_0100_ARTICLE,false)
Call SetTextbox("Sales Org\.","RMMW1-VKORG","",DT_MM42_0100_SALES_ORG,false)
Call SetTextbox("Distr\. Channel","RMMW1-VTWEG","",DT_MM42_0100_DISTR_CHANNEL,false)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",1,False)
Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",3,False)

Call PressEnter()     
Call TakeScreenShot
Call PressEnter()   
Call TakeScreenShot
Call ClickButton("Go to additional data   \(Ctrl\+F6\)", False)
Call SelectTab("TABSPR1","Vendor Data",False)
Call TakeScreenShot

Call PressEnter()
Call TakeScreenShot

 @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf5.xml_;_
'SAPGuiSession("Session").SAPGuiWindow("Change Article").InsightObject("InsightObject").Click @@ hightlight id_;_12_;_script infofile_;_ZIP::ssf6.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Change Article").InsightObject("InsightObject").Click
Wait 5
Call SetGridData("",1,"MEINH",DT_MM42_9901_GRIDCELL_0_AUN,False)
Call SetGridData("",1,"Packaging mat. type","ZPT1",False)
Call SetGridData("",1,"PM_MATNR",DT_MM42_9901_GRIDCELL_0_PACKAGING_MATERIAL_OCC1,False)
Call TakeScreenShot
 @@ hightlight id_;_5_;_script infofile_;_ZIP::ssf8.xml_;_
SAPGuiSession("Session").SAPGuiWindow("Change Article").InsightObject("InsightObject").Click @@ hightlight id_;_12_;_script infofile_;_ZIP::ssf9.xml_;_
wait 5
Call SetGridData("",2,"MEINH",DT_MM42_9901_GRIDCELL_0_AUN,False)
Call SetGridData("",2,"Packaging mat. type",DT_MM42_9901_GRIDCELL_0_PKMTT_OCC1,False)
Call SetGridData("",2,"PM_MATNR",DT_MM42_9901_GRIDCELL_0_PACKAGING_MATERIAL_OCC1,False)
Call TakeScreenShot


Call ClickButton("Go to main data   \(Ctrl\+Shift\+F3\)", False)
Call ClickButton("Display item details", False)
Call TakeScreenShot
Call SelectTab("INFOEXT","Sales 1",False)
Call TakeScreenShot
Call SetTableData("SAPLZMDAM_EXT_SCRTC_ECOMM","Consumption Units Per Article",1,"","",DT_MM42_2008_TABLECELL_CONSUMPTION_UNITS_PER_ARTICLE_0,False)
Call SetTableData("SAPLZMDAM_EXT_SCRTC_ECOMM","Net Content Per Consumption Unit",1,"","",DT_MM42_2008_TABLECELL_NET_CONTENT_PER_CONSUMPTION_UNIT_0,False)
Call SetTableData("SAPLZMDAM_EXT_SCRTC_ECOMM","Content Unit Consump",1,"","",DT_MM42_2008_TABLECELL_CONTENT_UNIT_CONSUMP_0,False)
Call PressEnter()
Call TakeScreenShot

Call SelectTab("INFOEXT","Logistics: DC/Store",False)
Call TakeScreenShot


Call SelectTab("INFOEXT","Store",False)
Call TakeScreenShot
Call SetTextbox("Label type","ZMDAM_MARAEXT-LABEL_TYPE","",DT_MM42_2006_ZMDAM_MARAEXTLABEL_TYPE,false)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR)

Call SelectRowGuiTableByRow("SAPLMGMWTAB_CONT_0100",1,False)
Call PressEnter()     
Call TakeScreenShot
''Call SetTextbox("X-site status","MARA-MSTAE","",DT_MM42_2004_XSITE_STATUS,false)
'Call SetTextboxNoLabel("MARA-MSTAE","",DT_MM42_2004_XSITE_STATUS,false)
Call ClickButton("Save   \(Ctrl\+S\)", False)
wait 4
Call PressEnter() 
Call VerifyStatusBar(DT_MM42_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)

 
Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


