
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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


gstrTestCaseName = "Test_ME21N_Create_PO_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\DLL\TASE_DT_ME21N_Create_PO.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
Call TakeScreenShot

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
''Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD",0,DT_ME21N_1105_VENDOR,False)  
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD",0,DT_ME21N_1105_VENDOR,False)  
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("Purch. Org.","MEPO1222-EKORG",0,DT_ME21N_1221_PURCH_ORG,False)     ' - Line (3)
Call SetTextbox("Purch. Group","MEPO1222-EKGRP",0,DT_ME21N_1221_PURCH_GROUP,False)     ' - Line (4)
Call SetTextbox("Company Code","MEPO1222-BUKRS",0,DT_ME21N_1221_COMPANY_CODE,False)     ' - Line (5)
   
'Call PressEnter()     ' - Line (7)
Call TakeScreenShot
Call SetTableData("SAPLMEGUITC_1211","Itm","1","","",DT_ME21N_1211_TABLECELL_ITM_0,False)     ' - Line (8)
Call SetTableData("SAPLMEGUITC_1211","A","1","","",DT_ME21N_1211_TABLECELL_A_0,False)     ' - Line (9)
Call SetTableData("SAPLMEGUITC_1211","Short Text","1","","",DT_ME21N_1211_TABLECELL_SHORT_TEXT_0,False)     ' - Line (10)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)     ' - Line (11)
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_ME21N_1211_TABLECELL_OUN_0,False)     ' - Line (12)
Call SetTableData("SAPLMEGUITC_1211","C","1","","",DT_ME21N_1211_TABLECELL_C_0,False)     ' - Line (13)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",DT_ME21N_1211_TABLECELL_DELIV_DATE_0,False)     ' - Line (14)
Call SetTableData("SAPLMEGUITC_1211","Net Price","1","","",DT_ME21N_1211_TABLECELL_NET_PRICE_0,False)  
Call SetTableData("SAPLMEGUITC_1211","Currency","1","","",DT_ME21N_1211_TABLECELL_CURRENCY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Per","1","","",DT_ME21N_1211_TABLECELL_PER_0,False)     ' - Line (16)
Call SetTableData("SAPLMEGUITC_1211","OPU","1","","",DT_ME21N_1211_TABLECELL_OPU_0,False)     ' - Line (17)
Call SetTableData("SAPLMEGUITC_1211","Matl Group","1","","",DT_ME21N_1211_TABLECELL_MDSE_CAT_0,False)     ' - Line (18)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)     ' - Line (19)
'Call SelectCellGuiTable("SAPLMEGUITC_1211","Plnt",refColumnName,refCellVal,False)       'Please provide reference column name and value     ' - Line (20)
Call TakeScreenShot
Call PressEnter()     ' - Line (21)
Call SetFocusGuiLabel("20010501",11,72,False)   
Call PressEnter()

Call SetComboByKey("MEACCT1200-KNTTP",DT_ME21N_1200_ACCASSCAT)     ' - Line (28)
Call SetTextbox("G/L Account","MEACCT1100-SAKTO","",DT_GL_ACCOUNT,False)     ' - Line (30)
Call SetTextbox("Business Area","COBL-GSBER","",DT_ME21N_1101_BUSINESS_AREA,False)     ' - Line (31)
Call SetTextbox("Cost Center","COBL-KOSTL","",DT_ME21N_1101_COST_CENTER,False)     ' - Line (32)
'Call FocusTextBox("Cost Center","COBL-KOSTL",False)     ' - Line (33)
Call TakeScreenShot
Call ClickButton("btn\[11\]",False)     ' - Line (34)
Call ClickButton("SPOP-VAROPTION1",True)     ' - Line (35)
Call GetStatusBar("item2","DT_DOC_NUM_OUTPUT")
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()








'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [30,110001]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

