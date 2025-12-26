
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_REB_07_01-ZMU8 - Enter CN Contract in SAP_P3_Post GR   
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



gstrTestCaseName = "Test_REB_07_01-ZMU8_P3_Post GR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_REB_07_01-ZMU8 - Enter CN Contract  in SAP_P3_Post GR.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''-----------------------login-----------------''''

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
''
''''''''-------Transaction Code MIGO-------'''''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
Call SetComboByKey("GODYNPRO-ACTION", DT_SETKEY_1)
Call SetComboByKey("GODYNPRO-REFDOC", DT_SETKEY_2)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2000_GODYNPROPO_NUMBER, False)
Call PressEnter()

Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call PressEnter()
Call TakeScreenShot
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",false)
If SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguiedit("guicomponenttype:=31","name:=BKPF-BELNR","attachedtext:=Document Number","Index:=0").Exist(1) Then
	wait 1
Else
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
End If
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call TakeScreenShot

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

