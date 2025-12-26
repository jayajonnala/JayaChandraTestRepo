
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :TASE_O2C_04_01_034-Claim process to MI stores - RW04 Dry Goods - Popesti_p2_Check_PO.xls     
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



gstrTestCaseName = "TASE_O2C_04_01_034-RW04 Dry Goods - Popesti_p2_Check_PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_O2C_04_01_034-Claim process to MI stores - RW04 Dry Goods - Popesti_p2_Check_PO.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''--------TransactionCode-WE02----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Created On","CREDAT-LOW","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
Call SetTextbox("to","CREDAT-HIGH","",ConvertDate(DT_WE02_1100_TO),False)
Call SetTextbox("Message Variant","MESCOD-LOW","",DT_WE02_1100_MESSAGE_VARIANT,False)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_WE02_1100_MESSAGE_FUNCTION,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
If SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Current Status","name:=EDIDC-STATUS","Index:=0").Exist(5)=False Then 
Call SelectColumnGuiGrid("Selected IDocs", 0,  "Created at", False)
Call ClickButtonToolbar("&SORT_DSC",0)
wait 5
Call SelectRowGuiGridbyRowNo("Selected IDocs", 0,1,False)
'Call DoubleClick()
Call DoubleClickGuiGridCell("Selected IDocs", 0, 1, "IDoc number", False)
End If
Call GetWindowValue("DT_IDOC_OUTPUT", False)
Call SelectMenuBar("Goto;Display Links")
Call ActivateCellGuiGridByRefVal("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", "Description", True)
Call SelectRowGuiGrid("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", True)
Call DoubleClickGuiGridCell("Relationships to IDoc:.*", 0, 4, "Document type", True)
Call TakeScreenShot()
Call GetTextboxValue("MEPO_TOPLINE-EBELN", "", "DT_WE02_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT", False)
Call VerifyTextBoxContent("Doc. date", "MEPO_TOPLINE-BEDAT", "", ConvertDate(DT_WE02_1105_CHECK_TEXT_OF_DOC_DATE), False)
Call ClickButton("btn\[21\]",False)
Call VerifyTableCellContent(2, "Status", "SAPDV70ATC_NAST3", DT_CHECK_TOOLTIP)
Call VerifyTableCellContent(2, "Output Type", "SAPDV70ATC_NAST3", DT_CHECK_OUTPUT)

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



