


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

gstrTestCaseName = "Test_07.1.4.3.01 AB Store Self Consumption _Fuels_P1_C1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_07.1.4.3.01 AB Store Self Consumption _Fuels_P1_C1E_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenSHot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------/POSDW/MON0 -----------------------------------
Call SetTextbox("Old Store","STOREO","",DT_ZTXCOPY_1000_OLD_STORE,False) 
Call SetTextbox("Old Posting Date","DAYO","",ConvertDate(DT_ZTXCOPY_1000_OLD_POSTING_DATE),False) 
'Call SetTextbox("Business transaction","BTY","",4,False) 
Call SetTextbox("Transaction Index","TIX-LOW","",DT_ZTXCOPY_1000_TO,False) 
Call SetTextbox("to","TIX-HIGH","",DT_ZTXCOPY_1000_TO,False) 
Call SetTextbox("New Store","STOREN","",DT_ZTXCOPY_1000_NEW_STORE,False) 
Call SetTextbox("New Posting Date","DAYN","",ConvertDAte(DT_ZTXCOPY_1000_NEW_POSTING_DATE),False)
Call TakeScreenSHot()
call SelectCheckbox("P_CHECK",0,"OFF",false)
call SelectCheckbox("LOG_SAV",0,"ON",false)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
call ClickButton("Back   \(F3\)",fALSE)
Call TakeScreenSHot()
call ClickButton("Back   \(F3\)",fALSE)
Call TakeScreenSHot()

Call ActivateNodeGuiTree(0,"User Menu for AUTOMI;W-POS: POS Workbench;POS WB: Selection and Overview")
Call TakeScreenSHot()
Call SetTextbox("Store","SO_STORE-LOW","",DT_ZTXCOPY_0201_STORE,False) 
Call SetTextbox("Posting Date","SO_DATE-LOW","",ConvertDate(DT_ZTXCOPY_0301_POSTING_DATE),False) 
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()

call GetGridContent("",1,"Transaction Number",1,"Processed",DT_PROCESSED,"DT_REFERENCE_OUTPUT")
'call ActivateItemGuiTree(1,"Inbound Monitor;03.10.2020;G061;Financial Transactions","Financial Transactions")
'Call SetTextbox("Store","DAYN","",DT_ZTXCOPY_0201_STORE,False)
'Call SetTextbox("Posting Date","DAYN","",DT_ZTXCOPY_0301_POSTING_DATE,False)
'
'call ClickButton("Execute   \(F8\)",fALSE)

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

Call SelectRowGuiGridbyRowNo("",0,1,False)

Call Click204ButtonToolBar("TRL_PROC",0)
Call TakeScreenSHot()

Call ClickButtonToolBar("POPUP_EXE",0)
Call TakeScreenSHot()
Call Click204ButtonToolBar("ERR_CLOSE",0)
Call SelectRowGuiGridbyRowNo("",0,1,False)
Call Click204ButtonToolBar("TRL_EDIT",0)
Call TakeScreenSHot()

Call ActivateNodeGuiTree(2,"#2;#4")
''call ClickButton("\Process Tasks Online",fALSE)
''Call TakeScreenSHot()
''call VerifyGridCellContent("",0,"STATUS",0,DT_ZTXCOPY_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON)
call GetGridContent("",0,"Quantity in Sales UOM",1,"Sales","EA","DT_SALES_UOM_OUTPUT")


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


