
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_016-Purchasing  - Corrections - Cancell- revers_P2_MAP
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_016-Purchasing  - Corrections - Cancell- revers_P2_MAP
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)






gstrTestCaseName = "Test_P2P_01_01_016- Cancell- revers_P2_MAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\DT_P2P_01_01_016-Purchasing  - Corrections - Cancell- revers_P2_MAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM) 
Call Login(DT_SAPUSER, DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (2)
Call PressEnter()

Call ClickButton("Display/Change   \(F7\)",False)

'Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME23N_0003_PUR_ORDER)
Call ClickButton("btn\[17\]",False)

Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur\. Order",False)
Call SetTextBox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_ME23N_0003_PUR_ORDER_OCC1,False)

Call PressEnter()
wait 5
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Texts",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)

Call TakeScreenShot()



Call TakeScreenShot()


Call FindRowNumber("","Short Text","RE-L","DT_ROW_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call GetGridContentByRefColumn("",0,"Movement type","102","Article Document","DT_GRIDCELL_0_BELNR_OUTPUT")
Call GetGridContentByRefColumn("",0,"Movement type","101","Article Document","DT_GRIDCELL_1_BELNR_OUTPUT")
Call GetGridContent("",0,"Article Document",DT_ROW_OUTPUT,"Short Text","RE-L","DT_GRIDCELL_3_BELNR_OUTPUT")
Call GetGridContent("",0,"Article Document",DT_ROW_OCC1,"Short Text","RE-L","DT_GRIDCELL_4_BELNR_OUTPUT")





Call SetTcode(DT_ME23N_0014_OKCD)   
Call PressEnter()
Call SetTextBox("Material","P_MATNR",0,DT_ME23N_1000_MATERIAL,False)
Call SetTextBox("Valuation Area","P_BWKEY",0,DT_ME23N_1000_VALUATION_AREA,False)

Call SetTextBox("Entry Date","R_DATUM-LOW",0,ConvertDate(DT_ME23N_1000_ENTRY_DATE),False)

Call TakeScreenShot()
Call PressEnter()

Call TakeScreenShot()
Call ClickButton("btn\[8\]",False)

Call TakeScreenShot()


Call ClickButton("btn\[29\]",False)
Call ClickButton("APP_WL_SING",False)

Call ClickButton("600_BUTTON",False)
Call CLickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot()
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_0,False)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_1,False)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_2,False)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_3,False)

Call ClickButton("btn\[0\]",False)

Call ClickButton("btn\[8\]",False)
Call TakeScreenShot()

Call ClickButton("btn\[0\]",False)

Call TakeScreenShot()


Call ClickButton("btn\[3\]",False)

Call SetTextBox("Material","P_MATNR",0,DT_ME23N_1000_MATERIAL_OCC1,False)
Call SetTextBox("Valuation Area","P_BWKEY",0,DT_ME23N_1000_VALUATION_AREA,False)

Call SetTextBox("Entry Date","R_DATUM-LOW",0,ConvertDate(DT_ME23N_1000_ENTRY_DATE),False)

Call TakeScreenShot()
Call PressEnter()

Call TakeScreenShot()
Call ClickButton("btn\[8\]",False)

Call TakeScreenShot()


Call ClickButton("btn\[29\]",False)
Call ClickButton("APP_WL_SING",False)

Call ClickButton("600_BUTTON",False)
Call CLickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot()
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_0,False)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_1,False)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_2,False)
Call SetTableData("SAPLALDBSINGLE","Single value","4","","",DT_ME23N_3010_TABLECELL_SINGLE_VALUE_3,False)

Call ClickButton("btn\[0\]",False)

Call ClickButton("btn\[8\]",False)
Call TakeScreenShot()

Call ClickButton("btn\[0\]",False)

Call TakeScreenShot()


Call LogOff()

Call FinalStatus()





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




