
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_054-EDI process in RW04 for goods receipt_P6_IDOC_POST   
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



gstrTestCaseName = "Test_P2P_01_01_054-EDI process P6_IDOC_POST"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_054-EDI process in RW04 for goods receipt_P6_IDOC_POST.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'//TransactionCode WE05//
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE05_1100_CHANGED_ON_DATE),false)   
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE05_1100_MESSAGE_VARIANT,false)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE05_1100_MESSAGE_FUNCTION,false)
Call SetTextbox("IDoc Number","DOCNUM-LOW","",DT_WE05_0100_GRIDCELL_0_IDOC_NUMBER,false)
Call SelectTab("TABSTRIP_IDOCTABBL","EDI",False)
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_INTERCHANGE_FILE_NO,false)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetWindowValue("DT_WE05_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1", false)

Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", "", DT_WE05_0100_CHECK_TEXT_OF_CURRENT_STATUS, false)

'''------/nBD87------------''''


Call SetTcode(DT_WE05_0100_OKCD)    
Call PressEnter()
Call TakeScreenShot

Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)


Call SetTextbox("IDoc Number","SX_DOCNU-LOW","",DT_WE05_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OCC1,false)
Call SetTextbox("Changed On","SX_UPDDA-LOW","",ConvertDate(DT_WE05_1100_CHANGED_ON_DATE_OCC1),False)      
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot


'''Wait till Application document posted text appears on tree. This is batch process and the status will reflect within 1 hour
Call ActivateNodeGuiTree(0, "#2;#1;#1")

''Call ActivateItemGuiTree("","Retail Pre-Production;IDoc in inbound processing;Application document posted","DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DESCRP")

Call GetGridContent("IDoc Selection", "", "Status text", "1", "Message Variant", "MII", "DT_WE05_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STAT_TXT_OUTPUT")

'''------/nMIR4------------''''
Call SetTcode(DT_WE05_0100_OKCD_OCC1)    
Call PressEnter()
Call TakeScreenShot

Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global", DataRowSet)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_WE05_6150_INVOICE_DOCUMENT_NO,false)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_WE05_6150_FISCAL_YEAR,false)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Follow-On Documents ...   \(F8\)",False)
Call SelectRowGuiGridbyRowNo("Documents in Accounting","",1,True)
Call GetGridContent("Documents in Accounting","","Doc. Number","1","Object Type Text","Accounting Document","DT_WE05_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER")
Call ClickButtonIfExist("Display Document   \(F2\)",True)

Call TakeScreenShot
Call GetTextboxValue("BKPF-BELNR", "", "DT_WE05_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT", False)
Call VerifyGridCellContent("", 1, "Posting Key", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "Posting Key", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 1, "Account", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "Account", "", DT_WE05_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

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



