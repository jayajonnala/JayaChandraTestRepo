

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.09.01 Proceeds from sale of gift cards with cash_Movement Type F506_Retail_C1E
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

gstrTestCaseName = "Test_09.06.05.09.01_F506_Retail_C1E"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.06.05.09.01 Proceeds from sale of gift cards with cash_Movement Type F506_Retail_C1E_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------/POSDW/MON0 -----------------------------------
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)

Call SelectTab("TABSTRIP_WB_TAB","Stores",False)
Call SetTextbox("Store","SO_STORE-LOW","",DT_POSDWMON1_0201_STORE,False) 
Call SetTextbox("Posting Date","SO_DATE-LOW","",ConvertDAte(DT_POSDWMON1_0301_POSTING_DATE),False) 
Call TakeScreenSHot()
Call SelectTab("TABSTRIP_WB_TAB2","Header",False)
Call TakeScreenSHot()
Call SetTextbox("Transaction Type","SO_TTCOD-LOW","",DT_POSDWMON1_0302_TRANSACTION_TYPE,False)
Call TakeScreenSHot()
call ClickButton("Execute   \(F8\)",fALSE)
Call TakeScreenSHot()
Call ActivateItemGuiTree(1, DT_ITEM_PATH, DT_ITEM_TEXT)
Call TakeScreenSHot()
call VerifyGridCellContent("",1,"Transaction Type",0,DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE)
call GetGridContent("",1,"Sales Totals Per Transaction",1,"Transaction Type",DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE,"DT_TURNOVER_OUTPUT")
call GetGridContent("",1,"Transaction Number",1,"Transaction Type",DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE,"DT_REFERENCE_OUTPUT")

'Added below steps to complete the approval process in C1E - Defect 37479
'Retrieve the transaction status from the grid
Call GetGridContentByRefColumn("", 1, "Transaction Type", DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TRANSTYPECODE,"Transaction Status", "DT_POSDWMON1_1000_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TR_STATUS_ICON_OUTPUT" )
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


