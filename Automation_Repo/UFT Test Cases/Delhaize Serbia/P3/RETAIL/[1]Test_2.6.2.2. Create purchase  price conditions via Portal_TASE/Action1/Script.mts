
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.2.2. Create purchase  price conditions via Portal
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.6.2.2. Create purchase  price conditions via Portal"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.7.1. Check Price Change File from Workflow.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME12 ----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Call SetTextbox("Vendor","EINA-LIFNR","",DT_ME12_0100_VENDOR,False)
Call SetTextboxNoLabel("EINA-LIFNR","",DT_ME12_0100_VENDOR,False)
Call SetTextbox("Article","EINA-MATNR","",DT_ME12_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org\.","EINE-EKORG","",DT_ME12_0100_PURCHASING_ORG,False)
Call TakeScreenShot()
Call PressEnter()    
Call TakeScreenShot()

Call GetTextboxValue("EINA-INFNR",0, "DT_ME12_0101_CHECK_TEXT_OF_INFO_RECORD_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ME12_0101_CHECK_TEXT_OF_INFO_RECORD_OUTPUT",DT_ME12_0101_CHECK_TEXT_OF_INFO_RECORD)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Conditions   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("New validity period with reference   \(F8\)",True)
Call TakeScreenShot()

Call SetTableDataNoRef("SAPMV13ATCTRL_D0201","CnTy",2,DT_ME12_0201_TABLECELL_CNTY_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_D0201","CnTy",3,DT_ME12_0201_TABLECELL_CNTY_2,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_D0201","Amount",1,DT_ME12_0201_TABLECELL_AMOUNT_0,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_D0201","Amount",2,DT_ME12_0201_TABLECELL_AMOUNT_1,False)
Call SetTableDataNoRef("SAPMV13ATCTRL_D0201","Amount",3,DT_ME12_0201_TABLECELL_AMOUNT_2,False)

Call PressEnter()
Call TakeScreenShot()

Call VerifyTableCellContent(1, "CnTy", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_CNTY_0)
Call VerifyTableCellContent(1, "Amount", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
Call VerifyTableCellContent(1, "Unit", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_UNIT_0)
Call VerifyTableCellContent(1, "Per", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_PER_0)
Call VerifyTableCellContent(1, "UoM", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_UOM_0)

Call VerifyTableCellContent(2, "CnTy", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_CNTY_1)
Call VerifyTableCellContent(2, "Amount", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_1)

Call VerifyTableCellContent(3, "CnTy", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_CNTY_2)
Call VerifyTableCellContent(3, "Amount", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_2)
Call VerifyTableCellContent(3, "Unit", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_UNIT_2)
Call VerifyTableCellContent(3, "Per", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_PER_2)
Call VerifyTableCellContent(3, "UoM", "SAPMV13ATCTRL_D0201", DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_UOM_2)

Call SetTextbox("Valid From","RV13A-DATAB", "", ConvertDate(DT_ME12_0201_VALID_FROM),False)
Call SetTextbox("Valid to","RV13A-DATBI", "", ConvertDate(DT_ME12_0201_VALID_TO),False)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Enter   \(F5\)",True)
Call TakeScreenShot()
Call VerifyStatusBar(LCase(DT_ME12_0100_CHECK_TEXT_OF_STATUSBAR))
Call PressEnter()
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()
