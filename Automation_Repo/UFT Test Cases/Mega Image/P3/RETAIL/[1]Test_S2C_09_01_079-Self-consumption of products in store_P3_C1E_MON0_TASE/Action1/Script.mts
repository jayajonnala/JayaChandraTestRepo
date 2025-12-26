
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_09_01_079-Self-consumption of products in store_P3_C1E_MON0
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

gstrTestCaseName = "Test_S2C_09_01_079-Self-consumption of products in store_P3_C1E_MON0"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2C_09_01_079-Self-consumption of products in store_P3_C1E_MON0.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-/POSDW/MON0 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Store","SO_STORE-LOW","",DT_POSDWMON0_0201_STORE,False)
Call FocusTextBox("Posting Date", "SO_DATE-LOW", False)
Call SendKey("{F2}")
Wait(5)
Call SelectRowGuiGrid("Posting Date", 0, "Description", "Single Value", True)
Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("Posting Date","SO_DATE-LOW","",ConvertDate(DT_POSDWMON0_0301_POSTING_DATE),False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP_WB_TAB2", "Header", False)
Call SetTextbox("Transaction Number","SO_TRNR-LOW","",DT_POSDWMON0_0302_TRANSACTION_NUMBER,False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",false)
'Added below steps to complete the approval process in C1E - Defect 37479
'Retrieve the transaction status from the grid
Call GetGridContentByRefColumn("", 1, "Transaction Number",DT_POSDWMON0_0302_TRANSACTION_NUMBER,"Transaction Status", "DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON_OUTPUT" )
' Check if the approval process is incomplete
If DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON_OUTPUT <>  DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS Then
'' Execute the approval process
Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call Click204ButtonToolBar("TRL_PROC", 0)
Call TakeScreenShot()
Call ClickButtonToolBar("POPUP_EXE", 0)
Call TakeScreenShot()
Call Click204ButtonToolBar("ERR_CLOSE", 0)
Call VerifyGridCellContent("",1,"Transaction Status",0,DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS)
Else 
' If the approval process is already completed, just verify the transaction status
Call VerifyGridCellContent("",1,"Transaction Status",0,DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS)
End If

Call TakeScreenShot
Call SelectAllRowGuiGrid("", 0, False)
Call DoubleClickGuiGridCell("", 0, 1, "Sales Totals Per Transaction", False)
Call TakeScreenShot
Call ActivateNodeGuiTree(2,"#2")
Call TakeScreenShot
Call ActivateNodeGuiTree(2,"#2;#5")
Call TakeScreenShot
Call ActivateNodeGuiTree(2,"#2;#4")
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GMREASONCODE)
Call VerifyGridCellContent("", 2, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GMREASONCODE)
Call VerifyGridCellContent("", 3, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_2_GMREASONCODE)
Call VerifyGridCellContent("", 4, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_3_GMREASONCODE)
Call VerifyGridCellContent("", 5, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_4_GMREASONCODE)
Call VerifyGridCellContent("", 6, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_5_GMREASONCODE)
Call VerifyGridCellContent("", 7, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_6_GMREASONCODE)
Call VerifyGridCellContent("", 8, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_7_GMREASONCODE)
Call VerifyGridCellContent("", 9, "RETAILTYPECODE", 0, DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_8_GMREASONCODE)

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



