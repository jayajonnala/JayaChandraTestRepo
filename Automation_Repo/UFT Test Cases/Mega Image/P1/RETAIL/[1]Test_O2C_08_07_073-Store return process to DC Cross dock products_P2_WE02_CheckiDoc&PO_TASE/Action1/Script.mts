
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_073-Store return process to DC Cross dock products_P2_WE02_CheckiDoc&PO     
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


gstrTestCaseName = "Test_O2C_08_07_073-Store return process to DC Cross dock products_P2_WE02_CheckiDoc&PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_O2C_08_07_073-Store return process to DC Cross dock products_P2_WE02_CheckiDoc&PO.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''--------TransactionCode-WE02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_WE02_1100_TO_OCC1),False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE02_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE02_1100_MESSAGE_FUNCTION,False)
Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
Call SetTextbox("Transfer File Reference","REFINT-LOW","",DT_WE02_1300_TRANSFER_FILE_REFERENCE,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call VerifyTextBoxContent("Current Status","EDIDC-STATUS",0,DT_WE02_CHECK_CURRENT_STATUS,False)

Call GetWindowValue("DT_IDOC_OUTPUT", False)
Call SelectMenuBar("Goto;Display Links")
Call ActivateCellGuiGridByRefVal("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", "Description", True)
Call SelectRowGuiGrid("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", True)
Call DoubleClickGuiGridCell("Relationships to IDoc:.*", 0, 4, "Document type", True)
Call VerifyTextBoxContent("Supplying Site", "MEPO_TOPLINE-SUPERFIELD", "",DT_WE02_1105_CHECK_TEXT_OF_SUPPLYING_SITE, False)

Call GetTableCellData("SAPLMEGUITC_1211", "Reqmt No.", 1, "<NA>", "<NA>", "DT_WE02_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0_OUTPUT", False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2", False)
Call SelectTab("HEADER_DETAIL", "Texts", False)
Call SelectTab("HEADER_DETAIL", "Communication", False)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Salesperson", "MEPOCOMM-VERKF", "",DT_WE02_1225_CHECK_TEXT_OF_SALESPERSON, False)
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4", False)
Call SelectTab("ITEM_DETAIL", "Purchase Order History", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call ClickButton("APP_WL_SING", True)
Call ClickButton("600_BUTTON", True)
Call SetTextbox("Short Text","%%DYN001-LOW","",DT_FILTER,True)
Call ClickButton("btn\[0\]", True)
Call VerifyGridCellContent("", 1, "Short Text", 0, lcase(DT_WE02_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BEWTK))
Call GetGridContentByRefColumn("", 0,  "Short Text", "Lfs","BELNR", "DT_BELNR_OUTPUT")

Call WriteRunTimeDataToExcelGlobalSheet("DT_BELNR_OUTPUT",DT_WE02_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)




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



