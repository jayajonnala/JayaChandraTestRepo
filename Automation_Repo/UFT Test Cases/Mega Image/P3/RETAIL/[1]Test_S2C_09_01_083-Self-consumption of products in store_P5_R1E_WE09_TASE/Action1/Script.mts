
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_09_01_083-Self-consumption of products in store_P5_R1E_WE09
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

gstrTestCaseName = "Test_S2C_09_01_083-Self-consumption of products in store_P5_R1E_WE09"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

''gstrInputExcelFilePathAndName="C:\Users\rsara\Desktop\TASEWork\Data\TASE_DT_S2C_09_01_083-Self-consumption of products in store_P5_R1E_WE09.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
''DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()



''--------TransactionCode-WE09---------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call SetTextbox("Created On","CREDAT-LOW", "",ConvertDate(DT_WE09_1000_CREATED_ON), False)
Call SetTextbox("to","CREDAT-HIGH", "", ConvertDate(DT_WE09_1000_TO), False)
Call SetTextbox("Logical Message","MESTYP-LOW", "","WPUBON", False)

Call SetTextbox("Search in Segment \.\.\.","SEGMENT1", "", DT_WE09_1000_SEARCH_IN_SEGMENT_, False)
Call SetTextbox("Search in Field \.\.\.","FIELD1_1", "", DT_WE09_1000_SEARCH_IN_FIELD_, False)
Call SetTextbox("for Value \.\.\.","VALUE1_1", "", DT_PVS, False)


Call ClickButtonIfExist("Execute   \(F8\)", False)

Call TakeScreenShot
If SAPGuiSession("micclass:=SAPGuiSession","activewindow:=IDoc.*").SAPGuiWindow("micclass:=SAPGuiWindow","text:=IDoc.*").SAPGuiLabel("micclass:=SAPGuiLabel","content:=00000.*","index:=0").Exist(0)  Then
	SAPGuiSession("micclass:=SAPGuiSession","activewindow:=IDoc.*").SAPGuiWindow("micclass:=SAPGuiWindow","text:=IDoc.*").SAPGuiLabel("micclass:=SAPGuiLabel","content:=00000.*","index:=0").SetFocus
	SAPGuiSession("micclass:=SAPGuiSession","activewindow:=IDoc.*").SAPGuiWindow("micclass:=SAPGuiWindow","text:=IDoc.*").SendKey F2
End If
''Call GetLabelContentByRefLabel("IDoc number",0,-54,"DT_IDOC_OUTPUT",False)

''Call ClickLabel(DT_IDOC_OUTPUT,0,False)
Call VerifyTextBoxContent("Current Status", "EDIDC-STATUS", 0, DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS, False)

Call SelectMenuBar("Goto;Display Links")
Call TakeScreenShot
wait 1
Call SelectRowGuiGridbyRowNo("Relationships.*", 0, 5, True)
Call DoubleClickGuiGridCell("Relationships.*", 0, 5, "Description", True)

Call TakeScreenShot
Call ClickButtonIfExist("Accounting overview   \(Shift\+F4\)", False)
Call TakeScreenShot

Call SelectRowGuiGridbyRowNo("Documents in Accounting", 0, 1, True)
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "Object type text", True)
Call TakeScreenShot


Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "Account", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "Account", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call VerifyGridCellContent("", 3, "Posting Key", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 3, "Account", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)

Call VerifyGridCellContent("", 4, "Posting Key", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 4, "Account", 0, DT_WE09_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

Call VerifyTextBoxContent("Reference", "BKPF-XBLNR", 0, DT_PVS, False)


Call LogOff()

Call FinalStatus ()










'''//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
'''Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'''Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'''Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 
''
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''
''
'''// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''' ................NOTE: 
'''.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'''.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'''.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'''.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
''' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''

