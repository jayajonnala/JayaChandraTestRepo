
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_08_07_092-Wholesale with WMS delivery - Ready meal customer_P4_CheckIdoc_WE02   
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



gstrTestCaseName = "Test_O2C_08_07_092-Wholesale  customer_P4_CheckIdoc_WE02"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI-O2C\TASE_DT_O2C_08_07_092-Wholesale with WMS delivery - Ready meal customer_P4_CheckIdoc_WE02.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
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
Call TakeScreenShot
Call SelectTab("TABSTRIP_IDOCTABBL", "EDI", False)
''Call SetTextbox("Reference to Message Group","REFGRP-LOW","",DT_WE02_1300_REFERENCE_TO_MESSAGE_GROUP,False) ''Not Required
Call SetTextbox("Message Reference","REFMES-LOW","",DT_WE02_1300_MESSAGE_REFERENCE,False)
Call PressEnter() 
Call ClickButton("Execute   \(F8\)",False)
Call VerifyGridCellContent("Selected IDocs", 1, "Status Grouping", 0, lcase(DT_WE02_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUSICON))
Call VerifyGridCellContent("Selected IDocs", 2, "Status Grouping", 0, lcase(DT_WE02_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_STATUSICON))
Call TakeScreenShot()
Call SelectColumnGuiGrid("Selected IDocs", 0, "IDOCTP", False)
Call ClickButtonToolBar("&MB_FILTER", 0)
Call SetTextbox("Basic type","%%DYN001-LOW","",DT_WE02_1105_BASIC_TYPE,True)
Call ClickButton("Execute   \(Enter\)", True)
Call SelectRowGuiGrid("Selected IDocs", 0, "Message reference", DT_WE02_1300_MESSAGE_REFERENCE, False)
Call DoubleClickGuiGridCell("Selected IDocs", 0, 1, "IDoc number", False)
Call SelectMenuBar("Goto;Display Links")
Call GetGridContent("Relationships to IDoc:.*", 0, "Description (Partner)", 1, "<NA>", "<NA>", "DT_WE02_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PARTNER_OUTPUT")

Call LogOff()

Call FinalStatus ()







'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



