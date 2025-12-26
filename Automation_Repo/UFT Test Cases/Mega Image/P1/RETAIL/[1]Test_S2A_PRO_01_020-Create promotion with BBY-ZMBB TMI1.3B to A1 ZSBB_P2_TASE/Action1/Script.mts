
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0113-Number Range for manual GR done in SAP_P1_GR
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
'.................Test Script Name :Test_P2P_01_01_0113-Number Range for manual GR done in SAP_P1_GR
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


gstrTestCaseName = "Test_S2A_PRO_01_020- TMI1.3B to A1 ZSBB_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0113-Number Range for manual GR done in SAP_P1_GR.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''''''''''''''''''''' TCODE MIGO '''''''''''''''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()

Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION,False)
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Bonus Buy",False)
Call TakeScreenShot()
Call ClickButton("BBY_CREATE",False)
Call SetTextBox("Bonus buy profile","KONBBYH-BPROF",0,DT_WAK2_2000_BONUS_BUY_PROFILE,True)
Call SetTextBox("Bonus buy","RDM_S_BBY_WA_HEADER-BBYNR",0,DT_WAK2_2000_BONUS_BUY,True)
Call TakeScreenShot()
Call PressEnter()
Call ClearTextBox("RDM_S_BBY_WA_HEADER_UI-BBYTEXT")
Call SetTextBox("Bonus Buy","RDM_S_BBY_WA_HEADER_UI-BBYTEXT",0,DT_WAK2_0600_BONUS_BUY,False)
Call ClickButton("BBY_TEXTS",False)
Call SetGridData("", 1, "Bonus Buy Text", DT_WAK2_2000_GRIDCELL_0_BONUS_BUY_TEXT, True)
Call TakeScreenShot()
Call ClickButton("btn\[8\]",False)
Call TakeScreenShot()
Call ClearTextBox("RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_TOTAL")
Call ClearTextBox("RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_ITEM")
Call SetTextBox("Total Message ID","RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_TOTAL",0,DT_WAK2_0600_TOTAL_MESSAGE_ID_OCC1,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("PUSH_CLOSE_ORG",False)
Call SetMultipleGridData("", 1, 1,"Line Item Identifier", DT_WAK2_2300_GRIDCELL_0_LINE_ITEM_IDENTIFIER, False)
Call TakeScreenShot()
Call  ClickButtonToolBar("BBY_INS_ROW",1)
Call SetMultipleGridData("", 1, 2,"Line Item Identifier", DT_WAK2_2300_GRIDCELL_1_LINE_ITEM_IDENTIFIER, False)
Call SetMultipleGridData("", 1, 2,"Discount Quantity", DT_WAK2_2300_GRIDCELL_1_QUANTITY, False)
Call SetMultipleGridData("", 1, 2,"Sales Unit", DT_WAK2_2300_GRIDCELL_1_UNIT, False)
Call SetMultipleGridData("", 1, 2,"Discount Type", DT_WAK2_2300_GRIDCELL_1_DISCOUNT_TYPE, False)
Call PressEnter()
Call SetMultipleGridData("", 1, 2,"Discount Value", DT_WAK2_2300_GRIDCELL_1_VALUE, False)
Call TakeScreenShot()

Call ClickButton("btn\[8\]",False)

Call ClickButton("btn\[3\]",False)

Call TakeScreenShot()

Call PressEnter()

Call ClickButton("btn\[11\]",False)

Call VerifyStatusBar("Promotion "&DT_WAK2_1100_PROMOTION&" saved")

Call LogOff()
Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''










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




