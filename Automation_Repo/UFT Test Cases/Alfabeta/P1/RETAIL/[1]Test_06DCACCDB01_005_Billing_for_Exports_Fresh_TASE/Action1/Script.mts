

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCACCDB01_005_Billing_for_Exports_Fresh
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_06DCACCDB01_005_Billing_for_Exports_Fresh"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_06DCACCDB01_005_Billing_for_Exports_Fresh_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=4
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'--------------------------------VA01-----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetTextbox("Order Type","VBAK-AUART","",DT_VA01_0101_ORDER_TYPE,False)   
Call SetTextbox("Sales Organization","VBAK-VKORG","",DT_VA01_0101_SALES_ORGANIZATION,False) 
Call SetTextbox("Distribution Channel","VBAK-VTWEG","",DT_VA01_0101_DISTRIBUTION_CHANNEL,False) 
Call SetTextbox("Division","VBAK-SPART","",DT_VA01_0101_DIVISION,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Sold-To Party","KUAGV-KUNNR","",DT_VA01_4701_SOLDTO_PARTY,False) 
Call PressEnter() 
Call TakeScreenShot()
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Article","1","","",DT_VA01_4900_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Order Quantity","1","","",DT_VA01_4900_TABLECELL_ORDER_QUANTITY_0,False)
Call SetTableData("SAPMV45ATCTRL_U_ERF_AUFTRAG","Site","1","","",DT_VA01_4900_TABLECELL_SITE_0,False)
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter() 
Call SendKey("{F2}")
Call SelectTab("TAXI_TABSTRIP_ITEM","Shipping",False)
Call SetTextbox("Storage Loc\.","VBAP-LGORT","",DT_VA01_4452_STOR_LOCATION,False) 
Call TakeScreenShot()
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call GetStatusBar("item2","DT_SALES_ORDER_OUTPUT")
Call VerifyStatusBar("Sales Order Push "& DT_SALES_ORDER_OUTPUT & " has been saved")

Call ClickButton("Back   \(F3\)",False) 
Call ClickButton("Back   \(F3\)",False) 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

''--------------------------------VL01n-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_VA01_0100_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Shipping point","LIKP-VSTEL","",DT_VA01_4001_SHIPPING_POINT,False) 
Call SetTextbox("Selection date","LV50C-DATBI","",ConvertDate(DT_VA01_4001_SELECTION_DATE),False) 
Call SetTextbox("Order","LV50C-VBELN","",DT_VA01_4001_ORDER,False)
Call PressEnter() 
Call TakeScreenShot()
Call PressEnter() 

Call SendKey("{F2}")
Call TakeScreenShot()
Call SelectTab("TAXI_TABSTRIP_ITEM","Documentary Batch",False)
Call SetGridData("",1,"Documentary Batch - Batch No.",DT_VA01_0201_GRIDCELL_0_DB_NO,False)
Call SetGridData("",1,"Qty in unit of entry",DT_VA01_0201_GRIDCELL_0_QTY_IN_UN_OF_ENTRY,False)
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False) 

Call GetStatusBar("item2","DT_OUTB_NO_OUTPUT")
Call VerifyStatusBar("Outb. del.(Affiilia) "& DT_OUTB_NO_OUTPUT & " has been saved")

Call ClickButton("Back   \(F3\)",False) 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

'--------------------------------VF01-----------------------------
Call SetTcode(DT_VA01_0100_OKCD_OCC1) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document","1","","",DT_OUTB_NO_OUTPUT,False) 
Call PressEnter() 
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "& DT_DOC_NO_OUTPUT & " has been saved")


Call LogOff()
Call FinalStatus ()

'




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





