
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0109-Store Non-contractual Free goods_P2_Display_Print
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


gstrTestCaseName = "Test_P2P_01_01_0109- goods_P2_Display_Print"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0109-Store Non-contractual Free goods_P2_Display_Print.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''---------Login-------------''''''
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''''-------Transaction Code ME21N-------'''''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)
Call PressEnter() 
Call SetComboByKey("GODYNPRO-REFDOC", DT_ARTICLE_DOC_TYPE)
Call PressEnter()
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIGO_2010_GODYNPROMAT_DOC, False) 
Call PressEnter()
Call SelectTab("TS_GOHEAD", "General", False)
Call VerifyTextboxContent("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_CHECK_TEXT_OF_DELIVERY_NOTE,False)
Call SelectTab("TS_GOITEM", "Where", False)
Call VerifyTextboxContent("Movement type","GOITEM-BWART","",DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
Call TakeScreenShot
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call ClickButton("FI Documents",false)
Call SelectRowGuiGrid("Documents in Accounting", 0, "OTEXT", "Accounting document", True)
Call GetGridContent("Documents in Accounting", 0, "DOCNR", 1, "<NA>", "<NA>", "DT_ARTICLE_DOCUMENT_NUMBER_OUTPUT")
Call DoubleClickGuiGridCell("Documents in Accounting", 0, 1, "DOCNR", True)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 4, "KTONR", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)

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



