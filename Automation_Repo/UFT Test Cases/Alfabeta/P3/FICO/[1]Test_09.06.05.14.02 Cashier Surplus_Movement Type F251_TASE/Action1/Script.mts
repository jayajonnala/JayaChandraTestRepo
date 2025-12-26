		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.14.02 Cashier Surplus_Movement Type F251
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.06.05.14.02 Cashier Surplus_Movement Type F251_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)
DT_FAGLL03_1000_POSTING_DATE= "01.01."& CSTR(Year(Date)-8)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FAGLL03_1000_TO),False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",false)
Call ActivateNodeGuiTree("0","#4;#8")
Call SetSpecialTextbox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot

Call ClickButtonIfExist("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB65_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB65_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)

Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Text", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",false)
Call TakeScreenShot
Call SetSpecialTextbox("Text","%%DYN001-LOW","","ΠΛΕΟΝΑΣΜΑΤΑ ΤΑΜΙΩΝ",True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Customer", False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",false)
Call TakeScreenShot
Call ClickButton("%_%%DYN002_%_APP_%-VALU_PUSH",True)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FAGLL03_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FAGLL03_3010_TABLECELL_SINGLE_VALUE_1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_FAGLL03_3010_TABLECELL_SINGLE_VALUE_2,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot
'
Call GetGridContentByTitle("", 0, "Amount in local currency", 1, "DT_FAGLL03_0500_GRIDCELL_0_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Amount in local currency", 2, "DT_FAGLL03_0500_GRIDCELL_0_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC1_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Amount in local currency", 3, "DT_FAGLL03_0500_GRIDCELL_0_OCC2_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC2_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Amount in local currency", 4, "DT_FAGLL03_0500_GRIDCELL_0_OCC3_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC3_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC3)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Customer", 1, "DT_FAGLL03_0500_GRIDCELL_0_OCC4_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC4_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC4)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Customer", 2, "DT_FAGLL03_0500_GRIDCELL_0_OCC5_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC5_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC5)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Customer", 3, "DT_FAGLL03_0500_GRIDCELL_0_OCC6_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC6_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC6)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Customer", 4, "DT_FAGLL03_0500_GRIDCELL_0_OCC7_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC7_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC7)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Posting Key", 1, "DT_FAGLL03_0500_GRIDCELL_0_OCC8_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC8_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC8)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Posting Key", 2, "DT_FAGLL03_0500_GRIDCELL_0_OCC9_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC9_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC9)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Posting Key", 3, "DT_FAGLL03_0500_GRIDCELL_0_OCC10_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC10_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC10)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Posting Key", 4, "DT_FAGLL03_0500_GRIDCELL_0_OCC11_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC11_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC11)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call VerifyGridCellContent("", 1, "KUNNR", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KUNNR)
Call VerifyGridCellContent("", 2, "KUNNR", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KUNNR)
Call VerifyGridCellContent("", 3, "KUNNR", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KUNNR)
Call VerifyGridCellContent("", 4, "KUNNR", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KUNNR)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 4, "BSCHL", 0, DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)

Call LogOff'
Call FinalStatus()''

