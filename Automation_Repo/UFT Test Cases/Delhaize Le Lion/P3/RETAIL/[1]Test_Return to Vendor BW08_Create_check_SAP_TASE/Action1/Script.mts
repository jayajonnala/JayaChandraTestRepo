
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Return to Vendor BW08_Create_check_SAP_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_Return to Vendor BW08_Create_check_SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_PRE_Get_Art_Doc_with_Reference_MB51.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Call SetTextbox("Article","MATNR-LOW","0",DT_MB51_1000_ARTICLE,False)
Call SetTextboxNoLabel("MATNR-LOW","0",DT_MB51_1000_ARTICLE,False)
Call SetTextboxNoLabel("WERKS-LOW","0",DT_MB51_1000_SITE,False)
Call SetTextboxNoLabel("LGORT-LOW","0",DT_MB51_1000_STORAGE_LOCATION,False)
Call SetTextboxNoLabel("BUKRS-LOW","0",DT_MB51_1000_COMPANY_CODE,False)
'Call SetTextboxNoLabel("XBLNR-LOW","0",DT_MB51_1000_DOCUMENT_HEADER_TEXT,False)
Call SetTextboxNoLabel("CPUDT-LOW","0",ConvertDate(DT_MB51_1000_ENTRY_DATE),False)
Call SetTextboxNoLabel("BWART-LOW","0",DT_MB51_1105_MOVEMENT_TYPE,False)
'Call SetTextboxNoLabel("XBLNR-LOW","0",DT_MB51_1000_SITE,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'' GetLabelContentByRefLabel(refLabelContent, xDifferenceValue, yDifferenceValue, dataTableColumnName, blnIsItPopup)
Call GetLabelContentByRefLabel("PO",0,-48,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN_OUTPUT,False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyifGuiLabelExistsByRelativeid(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART,"wnd\[0\]/usr/lbl\[6,5\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR,"wnd\[0\]/usr/lbl\[1,4\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS,"wnd\[0\]/usr/lbl\[68,5\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MENGE,"wnd\[0\]/usr/lbl\[129,5\]")
'' SetHorizontalScrollBar(ScrollBarPosition, blnIsItPopup)
Call SetHorizontalScrollBar(150,false)
Call GetLabelContentByRefLabel("Art\. Doc\.",0,-48,DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OUTPUT,False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''if grid is being displayed,execute below line of code'''
''Call GetGridContent("", 0, "MBLNR", 1, "<NA>", "<NA>", "DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR")
''Call VerifyGridCellContent("", 1, "BWART", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWART)
''Call VerifyGridCellContent("", 1, "MATNR", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
''Call GetGridContent("", 0, "EBELN", 1, "<NA>", "<NA>", "DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EBELN")
''Call VerifyGridCellContent("", 1, "BUKRS", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUKRS)
''Call VerifyGridCellContent("", 1, "MENGE", 0, DT_MB51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MENGE)
''
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


