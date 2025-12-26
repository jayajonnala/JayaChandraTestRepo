
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_Get_2_Articles_Qty_Vendor_From_PO_ME23N      
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


gstrTestCaseName = "Test_PRE_Get_2_Articles_Qty_Vendor_From_PO_ME23N"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_PRE_Get_2_Articles_Qty_Vendor_From_PO_ME23N.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''''//TransactionCode  /nME23N  //
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)

Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER_OCC1,True) 
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True) 
Call ClickButton("Other Document   \(Enter\)",True)

Call ClickButtonIfExist("Expand Header Ctrl\+F2", false)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot

Call GetTableCellData("SAPLMEGUITC_1211", "Article", 1, "<NA>", "<NA>", "DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0_OUTPUT", False)
Call GetTableCellData("SAPLMEGUITC_1211", "Article", 2, "<NA>", "<NA>", "DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_1_OUTPUT", False)
Call GetTableCellData("SAPLMEGUITC_1211", "PO Quantity", 1, "<NA>", "<NA>", "DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0_OUTPUT", False)
Call GetTableCellData("SAPLMEGUITC_1211", "PO Quantity", 2, "<NA>", "<NA>", "DT_ME23N_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_1_OUTPUT", False)
Call VerifyTextBoxNoLabelContent("MEPO_TOPLINE-EBELN", "", DT_ME23N_0003_PUR_ORDER_OCC1, False)
Call TakeScreenShot

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



