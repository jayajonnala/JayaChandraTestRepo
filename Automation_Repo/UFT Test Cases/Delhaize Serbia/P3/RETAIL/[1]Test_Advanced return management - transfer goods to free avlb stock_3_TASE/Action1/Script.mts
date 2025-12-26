
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Advanced return management - transfer goods to free avlb stock_3
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

gstrTestCaseName = "Test_Advanced return management - transfer goods to free avlb stock_3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_01-Regular purchasing in RW04  dry goods  via ME21N - P&Z_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'-------------------------------------Login----------------------
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'------------Transaction Code MSR_INSPWH----------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call SelectRadioButton("P_WH_16","No status restriction",True)
Call SelectCheckBox("P_WH_AL2",0,DT_MSR_INSPWH_0110_SHOW_ALL_ITEMS_OF_SELECTED_DELIVERIES,True)
Call SetTextbox("Delivery","INBD2-LOW","",DT_MSR_INSPWH_0110_DELIVERY,True)
Call TakeScreenShot
Call ClickButtonIFExist("Execute   \(F8\)",True)
Call TakeScreenShot
Call SendKey("{F2}")
wait 2


Call SetTextbox("Inspection Code","MSR_S_INSP_UI_HDR-DEC_CODE","",DT_MSR_INSPWH_2001_INSPECTION_CODE,False)
Call SetCombo("Follow-Up Activity",DT_MSR_INSPWH_2001_FOLLOWUP_ACTIVITY)
Call TakeScreenShot
Call ClickButton("Save and Confim LFU   \(Ctrl\+F8\)",False)

Call VerifyStatusBar(DT_MSR_INSPWH_2000_CHECK_TEXT_OF_STATUSBAR)

Call LogOff'
Call FinalStatus()
