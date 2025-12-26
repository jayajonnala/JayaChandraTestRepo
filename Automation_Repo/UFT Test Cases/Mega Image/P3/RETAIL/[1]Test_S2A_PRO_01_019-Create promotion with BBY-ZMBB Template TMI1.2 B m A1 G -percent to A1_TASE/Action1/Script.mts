
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_019-Create promotion with BBY-ZMBB Template TMI1.2 B m A1 G -percent to A1 
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


gstrTestCaseName = "Test_S2A_PRO_01_019-Create promotion with BBY-ZMBB Template TMI1.2 B m A1 G -percent to A1 "
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRO_01_019-Create promotion with BBY-ZMBB Template TMI1.2 B m A1 G -percent to A1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()   

Call SetTextBox("Promotion","WAKHD-AKTNR",0,DT_WAK2_1100_PROMOTION,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Bonus Buy",False)
Call TakeScreenShot()
Call ClickButton("Bonus Buy Create",False)
Call SetTextBox("Bonus buy","RDM_S_BBY_WA_HEADER-BBYNR",0,DT_WAK2_2000_BONUS_BUY,True)
Call SetTextBox("Bonus buy profile","KONBBYH-BPROF",0,DT_WAK2_2000_BONUS_BUY_PROFILE,TRUE)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call SetMultipleGridData("", 1, 1, "IDENT", DT_WAK2_2300_GRIDCELL_0_LINE_ITEM_IDENTIFIER, False)
Call SetMultipleGridData("", 1, 1, "BTYPE_R", DT_WAK2_2300_GRIDCELL_0_DISCOUNT_TYPE, False)
Call SetMultipleGridData("", 1, 1, "Discount Value", DT_WAK2_2300_GRIDCELL_0_VALUE, False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call SetTextBox("Total Message ID","RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_TOTAL",0,"",False)
Call SetTextBox("Item Message ID","RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_ITEM",0,"",False)
Call PressEnter()
Call TakeScreenShot()
Call SetTextBox("Item Message ID","RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_ITEM",0,DT_WAK2_0600_ITEM_MESSAGE_ID_OCC1,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call PressEnter()
Call SelectRowGuiGrid("Bonus Buys - Overview", 0, "Status", "Planned", False)
Call TakeScreenShot()
Call ClickButton("Bonus Buys Activate",False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_WAK2_GET_PROMOTION_NO_OUTPUT")
Call VerifyStatusBar("Promotion "&DT_WAK2_GET_PROMOTION_NO_OUTPUT&" saved")
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




