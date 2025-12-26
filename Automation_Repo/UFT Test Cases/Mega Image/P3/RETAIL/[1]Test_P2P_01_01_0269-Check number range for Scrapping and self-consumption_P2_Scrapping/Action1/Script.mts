
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0269-Check number range for Scrapping and self-consumption_P2_Scrapping       
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


gstrTestCaseName = "Test_P2P_01_01_0269 self-consumption_P2_Scrapping"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0269-Check number range for Scrapping and self-consumption_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''''''--------TransactionCode-ZMDIM_DESTR_STOCK----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_ZMDIM_DESTR_STOCK_1000_POSTING_DATE),False)
Call SetTextbox("Location","P_ABLAD", "", DT_ZMDIM_DESTR_STOCK_1000_LOCATION, False)
Call SetTextbox("Reason for Movement","P_GRUND", "", DT_ZMDIM_DESTR_STOCK_1000_REASON_FOR_MOVEMENT, False)
Call SetTextbox("Article","S_MATNR-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_ARTICLE, False)
Call SetTextbox("Customer","P_KUNNR", "", DT_ZMDIM_DESTR_STOCK_1000_CUSTOMER, False)
Call SetTextbox("Movement type","P_BWART","",DT_ZMDIM_DESTR_STOCK_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Storage Location","P_LGORT", "", DT_ZMDIM_DESTR_STOCK_1000_STORAGE_LOCATION, False)
Call SetTextbox("Site","S_WERKS-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_SITE, False)
Call SetTextbox("Company Code","S_BUKRS-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_COMPANY_CODE, False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot

Call SetGridData("", 1, "MENGE", DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_0_QUANTITY, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call DoubleClickGuiGridCell("", 0, 1, "MENGE", False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot

Call SelectAllRowGuiGrid("", 0, False)
Call TakeScreenShot
'Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call ClickButton("Save Documents   \(F8\)", False)
Call TakeScreenShot
Call GetGridContent("", 0, "MBLNR", 1, "<NA>", "<NA>", "DT_ZMDIM_DESTR_STOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBL_OUTPUT")
Call ClickCellGuiGrid("", 0, "MBLNR", 1, "<NA>", "<NA>", False)
Call TakeScreenShot
'''''Call GetTextboxValue("RM07M-MTSNR", 0, "DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ART_SLIP_OUTPUT", False)
'
Call SelectTab("TS_GOHEAD","General",False)
Call TakeScreenShot
Call GetTextboxValue("GOHEAD-MTSNR",0,"DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ART_SLIP_OCC1_OUTPUT",False)


''''''--------TransactionCode-ZMDIM_DESTR_STOCK----------''''

Call SetTcode(DT_ZMDIM_DESTR_STOCK_0500_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_ZMDIM_DESTR_STOCK_1000_POSTING_DATE),False)
Call SetTextbox("Location","P_ABLAD", "", DT_ZMDIM_DESTR_STOCK_1000_LOCATION, False)
Call SetTextbox("Reason for Movement","P_GRUND", "", DT_ZMDIM_DESTR_STOCK_1000_REASON_FOR_MOVEMENT, False)
Call SetTextbox("Article","S_MATNR-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_ARTICLE, False)
Call SetTextbox("Customer","P_KUNNR", "", DT_ZMDIM_DESTR_STOCK_1000_CUSTOMER, False)
Call SetTextbox("Movement type","P_BWART","",DT_ZMDIM_DESTR_STOCK_1000_MOVEMENT_TYPE,False)
Call SetTextbox("Storage Location","P_LGORT", "", DT_ZMDIM_DESTR_STOCK_1000_STORAGE_LOCATION, False)
Call SetTextbox("Company Code","S_BUKRS-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_COMPANY_CODE, False)
Call SetTextbox("Site","S_WERKS-LOW", "", DT_ZMDIM_DESTR_STOCK_1000_SITE, False)
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)", False)
Call TakeScreenShot

Call SetGridData("", 1, "MENGE", DT_ZMDIM_DESTR_STOCK_0500_GRIDCELL_0_QUANTITY, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call DoubleClickGuiGridCell("", 0, 1, "MENGE", False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call SelectAllRowGuiGrid("", 0, False)
Call TakeScreenShot
'Call SelectRowGuiGridbyRowNo("", 0, 1, False)
Call ClickButton("Save Documents   \(F8\)", False)
Call TakeScreenShot
Call GetGridContent("", 0, "MBLNR", 1, "<NA>", "<NA>", "DT_ZMDIM_DESTR_STOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR_OCC1_OUTPUT")
Call ClickCellGuiGrid("", 0, "MBLNR", 1, "<NA>", "<NA>", False)
Call TakeScreenShot
'''Call GetTextboxValue("RM07M-MTSNR", 0, "DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ART_SLIP_OCC1_OUTPUT", False)

Call SelectTab("TS_GOHEAD","General",False)
Call TakeScreenShot
Call GetTextboxValue("GOHEAD-MTSNR",0,"DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ART_SLIP_OCC1",False)
' GetInputFromExcel(InputFilePath, sheetName, IterationIndex)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContent("Article Slip","GOHEAD-MTSNR","",DT_ZMDIM_DESTR_STOCK_0420_CHECK_TEXT_OF_ART_SLIP_OCC2,False)


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


