
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_024-Import from EU Vendors GR before Invoice_P4_MAP
      
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


gstrTestCaseName = "Test_P2P_01_01_024-invoice_P4_MAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_024-Import from EU Vendors GR before Invoice_P4_MAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Job Name","BTCH2170-JOBNAME","",DT_MIGO_2170_JOB_NAME,False)
Call SetTextbox("User Name","BTCH2170-USERNAME","",DT_MIGO_2170_USER_NAME,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SetTcode(DT_MIGO_3000_OKCD_OCC4)     
Call PressEnter() 

Call SelectRadioButtonIfExist("P_FROM",2,False)
Call SetTextbox("Material","P_MATNR","",DT_ZMDIM_MBMAPCHANGES_1000_MATERIAL,False)
Call SetTextbox("Valuation Area","P_BWKEY","",DT_ZMDIM_MBMAPCHANGES_1000_VALUATION_AREA,False)
Call SetTextbox("Year","R_BDATJ-LOW","",DT_ZMDIM_MBMAPCHANGES_1000_YEAR,False)
Call SetTextbox("Reference Document","P_DOCU","",DT_ZMDIM_MBMAPCHANGES_1000_REFERENCE_DOCUMENT,False)
Call TakeScreenShot

Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call PressEnter() 
Call ClickButton("Set filter   \(Ctrl\+F5\)",false)
Call SelectCellGuiGrid("Column Set",0,1,"Column Name",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Reference Document Number","%%DYN001-LOW","",DT_ZMDIM_MBMAPCHANGES_1105_REFERENCE_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call GetGridContent("",0,"Total Valuated Stock Before Posting",1,"Reference Document Number",DT_ZMDIM_MBMAPCHANGES_1105_REFERENCE_DOCUMENT_NUMBER,"DT_VALUATED_STOCK_BEFORE_POSTING_OUTPUT")
Call GetGridContent("",0,"Value of Total Stock Before Posting",1,"Reference Document Number",DT_ZMDIM_MBMAPCHANGES_1105_REFERENCE_DOCUMENT_NUMBER,"DT_VALUE_OF_STOCK_BEFORE_POSTING_OUTPUT")
Call GetGridContent("",0,"Quantity Posted",1,"Reference Document Number",DT_ZMDIM_MBMAPCHANGES_1105_REFERENCE_DOCUMENT_NUMBER,"DT_QUANTITY_POSTED_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyGridCellContent("", 1, "Old V-Price", 0,round( DT_OLD_VPRICE,2))
Call GetGridContent("",0,"New V-Price",1,"Reference Document Number",DT_ZMDIM_MBMAPCHANGES_1105_REFERENCE_DOCUMENT_NUMBER,"DT_NEW_VPRICE_OUTPUT")
Call VerifyGridCellContent("", 1, "New V-Price", 0, DT_NEW_VPRICE_OUTPUT)
Call VerifyGridCellContent("", 1, "Status", 0, DT_MB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR)





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


