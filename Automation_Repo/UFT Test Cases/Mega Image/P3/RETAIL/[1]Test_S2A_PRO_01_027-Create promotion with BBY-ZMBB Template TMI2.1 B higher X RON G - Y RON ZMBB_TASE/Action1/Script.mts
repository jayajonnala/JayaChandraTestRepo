
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2A_PRO_01_027-Create promotion with BBY-ZMBB Template TMI2.1 B higher X RON G - Y RON ZMBB  
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


gstrTestCaseName = "Test_S2A_PRO_01_027-higher X RON G - Y RON ZMBB"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2A_PRO_01_027-Create promotion with BBY-ZMBB Template TMI2.1 B  higher X RON G - Y RON ZMBB.xls"
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
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Themes",False)
Call TakeScreenShot()
Call SetTableData("SAPMWAKAUEB_ZTHE", "Theme", 2, "<NA>", "<NA>", DT_WAK2_8214_TABLECELL_THEME_1, False)
Call PressEnter()
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_UEBERSICHT","Bonus Buy",False)
Call TakeScreenShot()
Call ClickButton("Bonus Buy Create",False)
Call SetTextBox("Bonus buy profile","KONBBYH-BPROF",0,DT_WAK2_2000_BONUS_BUY_PROFILE,True)
Call FocusTextBox("Bonus buy","RDM_S_BBY_WA_HEADER-BBYNR", True)
Call SendKey("{F4}")
Call SetTextBox("Bonus buy text","G_SELFLD_TAB-LOW",0,DT_WAK2_0220_BONUS_BUY_TEXT,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonifExist("Copy   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call SetTextBox("Bonus Buy","RDM_S_BBY_WA_HEADER_UI-BBYTEXT",0,DT_WAK2_0600_BONUS_BUY,False)
Call PressEnter()
Call ClickButton("BBY_TEXTS",False)
Call SetGridData("", 1, "BBYTEXT", DT_WAK2_2000_GRIDCELL_0_BONUS_BUY_TEXT_OCC1, True)
Call PressEnter()
wait 5
Call ClickButtonIfExist("Confirm   \(F8\)",True)

'''Call SetComboByKey("Link Category", DT_WAK2_2200_LINK_CATEGORY)
Call SetCombo("RDM_S_BBY_WA_HEADER_UI-BUYCON","AND Link")
Call SetTextBox("Total Message ID","RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_TOTAL",0,DT_WAK2_0600_TOTAL_MESSAGE_ID,False)
Call SetTextBox("Item Message ID","RDM_S_BBY_WA_HEADER_UI-ZZ_MSGID_ITEM",0,"",False)
Call PressEnter()
Call SelectRowGuiGridbyRowNo("", 1, 1, False)
Call ClickButtonToolBar("BBY_DEL_ROW", 1)
Call ClickButtonToolBar("BBY_INS_ROW", 1)
Call SetMultipleGridData("", 1, 1, "IDENT", DT_WAK2_2400_GRIDCELL_0_LINE_ITEM_IDENTIFIER, False)
Call SetMultipleGridData("", 1, 1, "MQUAN", DT_WAK2_2400_GRIDCELL_0_QUANTITY, False)
Call PressEnter()
Call ClickButtonToolBar("BBY_INS_ROW", 1)
Call SetMultipleGridData("", 1, 2, "IDENT", DT_WAK2_2400_GRIDCELL_1_LINE_ITEM_IDENTIFIER, False)
Call SetMultipleGridData("", 1, 2, "MQUAN", DT_WAK2_2400_GRIDCELL_1_QUANTITY, False)
Call PressEnter()
Call SetTextBox("Discount Value","RDM_S_BBY_WA_REWARD_UI-KBETR",0,DT_WAK2_2400_DISCOUNT_VALUE,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButtonIfExist("Activate Bonus Buy   \(F8\)",False)
Call VerifyTextBoxContent("Status", "RDM_S_BBY_WA_HEADER_UI-STATUS_TXT", "", lcase(DT_WAK2_0600_CHECK_TEXT_OF_STATUS), False)
Call TakeScreenShot()
Call ClickButton("Back   \(F3\)",False)
Call GetGridContent("Bonus Buys - Overview", 0, "BBYNR", 1, "<NA>", "<NA>", "DT_WAK2_1100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BBYNR_OUTPUT")
Call TakeScreenShot()
Call VerifyGridCellContent("Bonus Buys - Overview", 1, "STATUSTXT", 0, lcase(DT_WAK2_1100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUSTXT))
Call ClickButton("btn\[0\]",False)
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_WAK2_1100_CHECK_TEXT_OF_STATUSBAR)
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




