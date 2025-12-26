'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Intracompany Store returns to DC - SW31_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany Store returns to DC - SW31_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\Retail\DT_Intracompany Store returns to DC - SW31_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0",(Cint(DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode MSR_INSPWH----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRadioButtonIfExist(DT_MSR_INSPWH_0110_NO_STATUS_RESTRICTION,"",True)
Call SelectCheckbox("P_WH_AL2",1,DT_MSR_INSPWH_0110_SHOW_ALL_ITEMS_OF_SELECTED_DELIVERIES,True)
Call SetTextbox("Delivery","INBD2-LOW","",DT_MSR_INSPWH_0110_DELIVERY,True)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call ActivateNodeGuiTree("",DT_MSR_INSPWH_0110_DELIVERY)
Call SetTextbox("Inspection Code","MSR_S_INSP_UI_HDR-DEC_CODE","",DT_MSR_INSPWH_2001_INSPECTION_CODE,False)
Call SetComboByKey("MSR_S_INSP_UI_HDR-FU_CODE",DT_MSR_INSPWH_2001_FOLLOWUP_ACTIVITY)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()    
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Inspection Status","MSR_S_INSP_UI_HDR-INSP_STATUS_TXT","",Lcase(DT_MSR_INSPWH_2001_CHECK_TEXT_OF_INSPECTION_STATUS),False)

Call ClickButton("Save and Confim LFU   \(Ctrl\+F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

'verify statusbar with datasheet feed
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_MSR_INSPWH_2000_CHECK_TEXT_OF_STATUSBAR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

