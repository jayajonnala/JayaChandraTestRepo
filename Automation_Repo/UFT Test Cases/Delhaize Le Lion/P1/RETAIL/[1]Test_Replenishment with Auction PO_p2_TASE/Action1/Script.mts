
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : [1]Test_Replenishment with Auction PO_p2
'.................Test Scenario: AT_P2P_Fresh_Replenishment via Workbench and Auction PO
'.................TCode: ME23N, MEK1
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

gstrTestCaseName = "Test_Replenishment with Auction PO_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''ensure no open session
''''Call CloseSessionsSAP()
'''Login to SAP System

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode ME23N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'''ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)
wait(2)

'''search PO no
Call ClickButton("Other Purchase Order   \(Shift\+F5\)", False)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_PUR_ORDER,True)
Call ClickButton("Other Document   \(Enter\)", False)
Call TakeScreenShot()

'''' GetTableCellData(tableName, columnName, rowNumber, refColumnName, refCellValue, dataTableColumnName, blnIsItPopup)
'''Call GetTableCellData("SAPLMEGUITC_1211","Net Price",1,"PO Quantity",180,"DT_ME23N_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT",False)
'''Call WriteRunTimeDataToExcelGlobalSheet("DT_ME23N_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0_OUTPUT",DT_ME23N_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call GetTableCellData("SAPLMEGUITC_1211","Net Price",1,"PO Quantity",180,"DT_ME23N_TABLECELL_AMOUNT_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME23N_TABLECELL_AMOUNT_OUTPUT",DT_ME23N_TABLECELL_AMOUNT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''''validate amount
'''Call VerifyTableCellContent(1,"Net Price","SAPLMEGUITC_1211",DT_ME23N_6201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1,"Net Price","SAPLMEGUITC_1211",DT_ME23N_TABLECELL_AMOUNT)


'''----------------------Tcode MEK1----------------------------

Call SetTcode(DT_ME23N_0014_OKCD) 
Call PressEnter()     ' 
'''Call CheckTCodeScreen(DT_ME23N_0014_OKCD)
Call TakeScreenShot()
Call SetTextbox("Condition Type","RV13A-KSCHL","",DT_ME23N_0100_CONDITION_TYPE,FALSE)
Call PressEnter()
Call TakeScreenShot()
Call SelectRadioButton("RV130-SELKZ","Article Info Record",False)
Call ClickButton("Choose   \(Enter\)", False)
Call TakeScreenShot()
'''Call SetTextbox("Vendor","KOMG-LIFNR","",DT_ME23N_1018_VENDOR,FALSE)
Call SetTextboxNoLabel("KOMG-LIFNR","",DT_ME23N_1018_VENDOR,FALSE)
Call SetTextbox("Article","KOMG-MATNR","",DT_ME23N_1018_ARTICLE,FALSE)
Call SetTextbox("Purch\. Organization","KOMG-EKORG","",DT_ME23N_1018_PURCH_ORGANIZATION,FALSE)
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Infotype",1, , ,DT_ME23N_1018_TABLECELL_Infotype_0,False)
'''Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1, , ,DT_ME23N_1018_TABLECELL_AMOUNT_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Amount",1, , ,DT_ME23N_TABLECELL_AMOUNT_1,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Unit",1, , ,DT_ME23N_1018_TABLECELL_UNIT_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","per",1, , ,DT_ME23N_1018_TABLECELL_PER_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","UoM",1, , ,DT_ME23N_1018_TABLECELL_UOM_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Calculat.type",1, , ,DT_ME23N_1018_TABLECELL_Calculattype_0,False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Valid From",1, , ,ConvertDate(DT_ME23N_1018_TABLECELL_ValidFrom_0),False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Valid To",1, , ,ConvertDate(DT_ME23N_1018_TABLECELL_ValidTo_0),False)
Call SetTableData("SAPMV13ATCTRL_FAST_ENTRY","Exclusion",1, , ,DT_ME23N_1018_TABLECELL_Exclude_0,False)
Call TakeScreenShot()
Call PressEnter()
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Enter   \(F5\)",True)
'''Call ValidateStatusBarMsg()
Call VerifyStatusBar(DT_ME23N_1018_CHECK_TEXT_OF_STATUSBAR)

Call LogOff()
Call FinalStatus()

'*********************************************End Of Script*********************************************************************

