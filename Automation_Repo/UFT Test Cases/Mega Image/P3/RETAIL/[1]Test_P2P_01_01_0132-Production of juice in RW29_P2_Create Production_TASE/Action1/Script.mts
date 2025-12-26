
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0132-Production of juice in RW29_P2_Create Production     
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


gstrTestCaseName = "Test_P2P_01_01_0132_P2_Create Production"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0132-Production of juice in RW29_P2_Create Production.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''''--------------login----------------'''''
''
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

''''--------TransactionCode-MB52----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Storage Location","LGORT-LOW","",DT_MB52_1000_STORAGE_LOCATION,False)
Call SetTextbox("Site","WERKS-LOW","",DT_MB52_1000_SITE,False)
Call SetTextbox("Article","MATNR-LOW","",DT_MB52_1000_ARTICLE,False)
Call TakeScreenShot
Call ClickButtonIfExist("%_MATNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", 2, "", "", DT_MB52_3010_TABLECELL_SINGLE_VALUE_0, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 3, "", "", DT_MB52_3010_TABLECELL_SINGLE_VALUE_1, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 4, "", "", DT_MB52_3010_TABLECELL_SINGLE_VALUE_2, True)
Call SetTableData("SAPLALDBSINGLE", "Single value", 5, "", "", DT_MB52_3010_TABLECELL_SINGLE_VALUE_3, True)
Call ClickButtonIfExist("Copy   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("PA_FLT", "Non-Hierarchical Representation", False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call GetLabelContentByRefLabel("Unrestr.", 0, -32, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT", False)
''Call GetLabelContentByRefLabel("Unrestr.", 0, -48, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT", False)
Call GetLabelContentByRefLabel("Unrestr.", 0, -48, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT", False)
''Call GetLabelContentByRefLabel("Unrestr.", 0, -72, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT", False)
Call GetLabelContentByRefLabel("Unrestr.", 0, -64, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUTPUT", False)
''Call GetLabelContentByRefLabel("Unrestr.", 0, -96, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUTPUT", False)
Call GetLabelContentByRefLabel("Unrestr.", 0, -80, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC3_OUTPUT", False)
''Call GetLabelContentByRefLabel("Unrestr.", 0, -120, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC3_OUTPUT", False)
Call GetLabelContentByRefLabel("Unrestr.", 0, -96, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC4_OUTPUT", False)
''Call GetLabelContentByRefLabel("Unrestr.", 0, -144, "DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OCC4_OUTPUT", False)
Call TakeScreenShot


''''--------TransactionCode-/nMD12----------''''

Call SetTcode(DT_MB52_0120_OKCD)     
Call PressEnter() 

Call SetTextbox("Planned Order","RM61P-PLNUM","",DT_MB52_0101_PLANNED_ORDER,False)
Call PressEnter()
Call PressEnter()
Call PressEnter()

Call ClickButton("Component Overview   \(F5\)",False)
Call TakeScreenShot

''Call VerifyTableCellContent(1, "Requirement Qty", "SAPLM61QTC_517", DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_0)
''Call VerifyTableCellContent(2, "Requirement Qty", "SAPLM61QTC_517", DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_1)
''Call VerifyTableCellContent(3, "Requirement Qty", "SAPLM61QTC_517", DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_2)
''Call VerifyTableCellContent(4, "Requirement Qty", "SAPLM61QTC_517",DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_3)

Call SelectRowGuiTableByRow("SAPLM61QTC_517",1,False)
Call ClickButton("Choose   \(F2\)",False)
Call VerifyTextBoxContent("Reqmt Qty","MDPME-ERFMG",0,DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_0,False)
Call SelectRowGuiTableByRow("SAPLM61QTC_517",2,False)
Call ClickButton("Choose   \(F2\)",False)
Call VerifyTextBoxContent("Reqmt Qty","MDPME-ERFMG",0,DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_1,False)
Call SelectRowGuiTableByRow("SAPLM61QTC_517",3,False)
Call ClickButton("Choose   \(F2\)",False)
Call VerifyTextBoxContent("Reqmt Qty","MDPME-ERFMG",0,DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_2,False)
Call SelectRowGuiTableByRow("SAPLM61QTC_517",4,False)
Call ClickButton("Choose   \(F2\)",False)
Call VerifyTextBoxContent("Reqmt Qty","MDPME-ERFMG",0,DT_MB52_0517_CHECK_TEXT_OF_TABLECELL_REQUIREMENT_QTY_3,False)
'
''
'''--------TransactionCode-/nmfbf----------''''

Call SetTcode(DT_MB52_0115_OKCD)     
Call PressEnter() 

Call SetTextbox("Planned order","RM61B-PLNUM","",DT_MB52_0801_PLANNED_ORDER,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("btn\[8\]",False) ''post with correction
Call TakeScreenShot
Call SetTableData("SAPLCOWBTCTRL_0540", "Quantity", 1, "<NA>", "<NA>", DT_MB52_0540_TABLECELL_QUANTITY_0, False)
Call SetTableData("SAPLCOWBTCTRL_0540", "Quantity", 2, "<NA>", "<NA>", DT_MB52_0540_TABLECELL_QUANTITY_1, False)
Call SetTableData("SAPLCOWBTCTRL_0540", "Quantity", 4, "<NA>", "<NA>", DT_MB52_0540_TABLECELL_QUANTITY_3, False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

'''For Iterator = 1 To 4 Step 1
''' Stausbartext = SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiStatusBar(guiStatusBar).GetROProperty("Text")
'''	If Instr(1,Stausbartext,"GR and GI with document") > 0 then 
'''		Exit For		
'''		Else
'''		Call PressEnter()
'''	End If
'''Next
''While SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=31","attachedtext:=Goods Receipt Quanti","name:=G_COWB_HEADER-MENGE","Index:=0").Exist(5)=True 
''Call PressEnter()  
''Wend
Call PressEnter()  
'Call PressEnter()  
Call GetTextStatusBar("DT_GET_PRODUCTION_MESSAGE_OUTPUT")
Call VerifyStatusBar(Lcase(DT_GET_PRODUCTION_MESSAGE_OUTPUT))

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



