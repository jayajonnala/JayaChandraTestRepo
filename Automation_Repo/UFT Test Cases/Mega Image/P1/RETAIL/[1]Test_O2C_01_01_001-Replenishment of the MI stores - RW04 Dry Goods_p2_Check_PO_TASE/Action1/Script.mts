
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_O2C_01_01_001-Replenishment of the MI stores - RW04 Dry Goods_p2_Check_PO      
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
gstrTestCaseName = "Test_O2C_01_01_001-Replenishment of the MI stores - RW04 Dry Goods_p2_Check_PO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\ssahoo\Desktop\TASEWork\Data\P1-MI\TASE_DT_O2C_01_01_001-Replenishment of the MI stores - RW04 Dry Goods_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
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
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

''''Call DoubleClickGuiGridCell("Selected IDocs", 0, 1, "IDoc number", False)
If SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Current Status","name:=EDIDC-STATUS","Index:=0").Exist(5)=False Then 
Call SelectRowGuiGridbyRowNo("Selected IDocs", 0,1,False)
Call DoubleClickGuiGridCell("Selected IDocs", 0, 1, "IDoc number", False)
End If

Call SelectMenuBar("Goto;Display Links")
Call TakeScreenShot()
Call ActivateCellGuiGridByRefVal("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", "Description", True)
Call SelectRowGuiGrid("Relationships to IDoc:.*", 0, "Document type", "Purchase Order", True)
Call DoubleClickGuiGridCell("Relationships to IDoc:.*", 0, 4, "Document type", True)
Call TakeScreenShot()
Call GetTextboxValue("MEPO_TOPLINE-EBELN", "", "DT_WE02_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT", False)
Call VerifyTextBoxContent("Doc. date", "MEPO_TOPLINE-BEDAT", "", ConvertDate(DT_WE02_1105_CHECK_TEXT_OF_DOC_DATE), False)
Call TakeScreenShot

''----------------------------------------------------------
'Call SetTcode("/nME22N")  
'Call TakeScreenShot
'Call PressEnter()     
'Call TakeScreenShot
'
'Call ClickButton("Other Purchase Order   \(Shift\+F5\)",false)
'Call TakeScreenShot
'Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_WE02_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT,True)
'Call TakeScreenShot
'Call ClickButton("Other Document   \(Enter\)",True)
'Call TakeScreenShot
'
'Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
'Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
'Call TakeScreenShot
'
'Call PressEnter()     
'wait 5
'Call SetTableData("SAPLMEGUITC_1320","Stat. Del. Dte","2","","",ConvertDate(DT_WE02_1100_CREATED_ON),False)
'Call TakeScreenShot
'Call ClickButton("Save   \(Ctrl\+S\)",False)
'Call TakeScreenShot
'Call ClickButton("SPOP-VAROPTION1",True)
'Call TakeScreenShot
'
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



