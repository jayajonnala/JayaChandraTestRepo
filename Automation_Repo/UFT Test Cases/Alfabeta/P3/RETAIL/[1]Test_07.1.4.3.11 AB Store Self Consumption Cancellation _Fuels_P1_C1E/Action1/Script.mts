

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_07.1.4.3.11 Cancellation _Fuels_P1_C1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.1.4.3.11 AB Store Self Consumption Cancellation _Fuels_P1_C1E.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'''//-----------------------------------/POSDW/MON0 -----------------------------------

Call SelectTab("TABSTRIP_WB_TAB","Stores",False)
Call SetTextbox("Store","SO_STORE-LOW","",DT_POSDWMON0_0201_STORE,False) 

Call SetTextbox("Posting Date","SO_DATE-LOW","",ConvertDate(DT_POSDWMON0_0301_POSTING_DATE),False) 
''Call SelectTab("TABSTRIP_WB_TAB2","Header",False)
''Call SetTextbox("Transaction Type","SO_TTCOD-LOW","",DT_POSDWMON1_0302_TRANSACTION_TYPE,False)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
Call SelectRowGuiGridbyRowNo("",0,1,False)
Call Click204ButtonToolBar("TRL_COPY",0)
Call TakeScreenSHot()
Call SelectRowGuiGridbyRowNo("",0,1,False)
Call Click204ButtonToolBar("TRL_EDIT",0)
Call TakeScreenSHot()
Call ActivateNodeGuiTree(2,"#2;#4")
Call TakeScreenSHot()

Call SetGridData("",1,"Quantity in Sales UOM",DT_POSDWMON0_1000_GRIDCELL_0_QUANTITY_SALES_UOM,False)
Call TakeScreenSHot()
Call ActivateNodeGuiTree(2,"#1")
Call SetGridData("",1,"POS Number",DT_POSDWMON0_1000_GRIDCELL_0_POS_NUMBER,False)

Call TakeScreenSHot()
call ClickButton("Save   \(Ctrl\+S\)",fALSE)
call ClickButton("Back   \(F3\)",fALSE)
Call SelectRowGuiGridbyRowNo("",0,1,False)

Call Click204ButtonToolBar("TRL_PROC",0)
Call TakeScreenSHot()

Call ClickButtonToolBar("POPUP_EXE",0)

Call TakeScreenSHot()
Call Click204ButtonToolBar("ERR_CLOSE",0)
call VerifyGridCellContent("",0,"STATUS",0,DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON)
call GetGridContent("",1,"Transaction Number",1,"STATUS",DT_POSDWMON0_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON,"DT_REFERENCE_OUTPUT")

'Added below steps to complete the approval process in C1E - Defect 37479
'Retrieve the transaction status from the grid
'Call GetGridContentByRefColumn("", 1, "Transaction Type", DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE,"Transaction Status", "DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON_OUTPUT" )
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

Call LogOff()
Call FinalStatus ()
'''
'''' ActivateItemGuiTree(treeIndex, itemPath, itemText)
''''call ActivateItemGuiTree(2,"Inbound Monitor;01.10.2020;G061;Financial Transactions","Financial Transactions")
'''' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
'''Call ActivateItemGuiTree(1,"Inbound Monitor;01.10.2020;G061;Financial Transactions","Financial Transactions")
'''' Click204ButtonToolBar(buttonName, toolbarIndex)
''''Click204ButtonToolBar
'''' ClickButtonToolBar(buttonName, toolbarIndex)
'''CALL ClickButtonToolBar("TRL_PROC",16)
'''CALL ClickButtonToolBar("POPUP_EXE",0)
'''
''''SAPGuiSession("Session").SAPGuiWindow("POS Workbench").Maximize
''''SAPGuiSession("Session").SAPGuiWindow("POS Workbench").SAPGuiGrid("GridViewCtrl").SelectCell 2,"Ready"
''''SAPGuiSession("Session").SAPGuiWindow("POS Workbench").SAPGuiToolbar("GridToolbar").PressButton "TRL_PROC"
''''Browser("Execution Report Results").Page("Execution Report Results").Sync
''''SAPGuiSession("Session").SAPGuiWindow("Process Tasks Online").Page("Page").Frame("Frame").WebList("INP_I_1_I_").Select "BW01 - Export to BW"
''''SAPGuiSession("Session").SAPGuiWindow("Process Tasks Online").Page("Page_2").Frame("Frame").WebList("INP_I_1_I_").Select "TBAL - Transaction Balance check"
''''Browser("Execution Report Results").Page("Execution Report Results").Image("Image_2019-04-22_10-22-28_760").Click
''''SAPGuiSession("Session").SAPGuiWindow("Process Tasks Online").SAPGuiToolbar("ToolBarControl").PressButton "POPUP_EXE"
'''''Call  ActivateNodeGuiTree(2,"Inbound Monitor;01.10.2020;G061;Financial Transactions")
''''call VerifyGridCellContent("",1,"Account",0,DT_FBL5N_0301_CHECK_TEXT_OF_CUSTOMER)
''''call VerifyGridCellContent("",1,"Transaction Type",0,DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE)
''''' ActivateNodeGuiTree(treeIndex, itemPath)
''''
'''''SAPGuiSession("Session").SAPGuiWindow("POS Workbench").Maximize
'''''SAPGuiSession("Session").SAPGuiWindow("POS Workbench").SAPGuiTree("TableTreeControl").ActivateItem "Inbound Monitor;01.10.2020;G061;Financial Transactions","Financial Transactions"
''''
''''



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


