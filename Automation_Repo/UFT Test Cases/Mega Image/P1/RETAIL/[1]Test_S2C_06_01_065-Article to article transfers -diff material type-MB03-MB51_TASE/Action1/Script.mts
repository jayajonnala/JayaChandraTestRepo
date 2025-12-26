

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_06_01_065-Article to article transfers -diff material type-MB03-MB51
      
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


gstrTestCaseName = "Test_S2C_06_01_065-diff material type-MB03-MB51"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2C_06_01_065-Article to article transfers -diff material type-MB03-MB51.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_MB03_0460_ARTICLE_DOC,False)
Call SetTextbox("Art\. Doc\. Year","RM07M-MJAHR","",DT_MB03_0460_ART_DOC_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Quantity","MSEG-ERFMG",0,DT_MB03_0420_CHECK_TEXT_OF_QUANTITY,False)
Call VerifyTextBoxContent("Quantity","MSEG-ERFMG",1,DT_MB03_0420_CHECK_TEXT_OF_QUANTITY_OCC1,False)
Call VerifyTextBoxContent("S","DM07M-VZBEW",0,DT_MB03_0420_CHECK_TEXT_OF_S,False)
Call VerifyTextBoxContent("S","DM07M-VZBEW",1,DT_MB03_0420_CHECK_TEXT_OF_S_OCC1,False)
Call VerifyTextBoxContent("Material Description","MSEG-MATNR",0,DT_MB03_0420_CHECK_TEXT_OF_ARTICLE_DESCRIPTION,False)
Call VerifyTextBoxContent("Material Description","MSEG-MATNR",1,DT_MB03_0420_CHECK_TEXT_OF_ARTICLE_DESCRIPTION_OCC1,False)
Call SetTcode(DT_MB03_0420_OKCD)     
Call PressEnter() 
Call SetTextbox("Company Code","BUKRS-LOW","",DT_MB03_1000_COMPANY_CODE,False)
Call SetTextbox("Article Document","MBLNR-LOW","",DT_MB03_1000_ARTICLE_DOCUMENT,False)
Call ClickButton("btn\[8\]",False)
Call ClickButton("Current Display Variant   \(Ctrl\+F8\)",False)
Call ClickButtonIfExist("Detail List   \(Ctrl\+Shift\+F12\)",False)
Call SelectCellGuiGrid("Column Set",0,20,"Column Name",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot
Call SelectColumnGuiGrid("",0,"Ext. Amount in Local Currency",False)
Call ClickButton("Add Up Values   \(Shift\+F7\)",False)
Call VerifyGridCellContent("", 2, "Ext. Amount in Local Currency", 0, DT_MB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMBTR)
Call TakeScreenShot
Call VerifyGridCellContent("", 3, "Quantity", 0, DT_MB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ERFMG)
Call DoubleClickGuiGridCell("", 0, 1, "Article Document", False)
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call TakeScreenShot
Call ClickButtonIfExist("FI Documents",False)
wait (3)
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_MB51_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

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


