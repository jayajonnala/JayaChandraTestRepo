
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06OM08_002_Create_Store_order_for_integrated_store_or_DC_Fresh
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCACCDB02_002_Create_Delivery_for_integrated_store_Fresh_via_SAP
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06DCACCDB02_002_Fresh_via_SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_06DCACCDB02_002_Create_Delivery_for_integrated_store_Fresh_via_SAP_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  VL10G------------------------------------------------
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_VL10G_1020_SALES_DOCUMENT,False) 
Call TakeScreenShot()  
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call SetTextboxNoLabel("P_LERUL","","",False)  
Call PressEnter() 
Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_VL10G_1030_PURCHASING_DOCUMENT,False) 
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()
Call ClickCellGuiGrid("",0,"Originating Document",1,"","",False)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Custom Fields",False)
Call VerifyTextBoxContent("Issuing Storage","CI_EKKODB-ZZ_RESLO","",DT_VL10G_9000_CHECK_TEXT_OF_ISSUING_STORAGE,False)
Call ClickButton("Back   \(F3\)",False) 


Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Call SelectRowGuiGrid("",0,"Originating Document",DT_VL10G_1030_PURCHASING_DOCUMENT,False)
Call ClickButton("Create Delivery in Background   \(Shift\+F7\)",False) 
Call TakeScreenShot()
Call SelectColumnGuiGrid("",0,"Sales Document",False)
Call ClickButton("Sort in Descending Order   \(Ctrl\+Shift\+F4\)",False)
Call GetGridContent("",0,"Sales Document",1,"<NA>","<NA>","DT_VL10G_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_VBELN_OUTPUT")
Call ClickCellGuiGrid("",0,"Sales Document",1,"","",False)


'--------------------------------------------  VL03N------------------------------------------------
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call ClickButton("Display <-> Change   \(Ctrl\+F1\)",False) 
Call TakeScreenShot()
Call VerifyTextBoxContent("Outbound deliv\.","LIKP-VBELN","",DT_VL10G_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_VBELN_OUTPUT,False)
Call SetTableData("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","1","","",DT_VL10G_1102_TABLECELL_DOCU_BATCH_NO_0,False)  
Call ClickButton("Save   \(Ctrl\+S\)",False) 

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call GetStatusBar("item2","DT_OUTBOUND_DEL_OUTPUT")
Call  VerifyStatusBar("Outb. del.(Intracy) "& DT_OUTBOUND_DEL_OUTPUT &" has been saved")

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


