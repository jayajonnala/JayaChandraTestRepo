		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.06.01 Cheques Receivables received in Store - Trasfer
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

gstrTestCaseName = "Test_09.05.01.06.01 Cheques Receivables received in Store - Trasfer"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-FTR03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectRadioButton("X8","Removed from Portfolio",False)
Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_FTR03_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_FTR03_1000_POSTING_DATE),False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR03_1000_COMPANY_CODE,False) 
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot
Call SetTextbox("Old Portfolio","P_PORT4","",DT_FTR03_1002_OLD_PORTFOLIO,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot


Call GetLabelContentByRefLabel("DocumentNo",0,-32,"DT_FTR03_0120_NO_NAME_INDEX_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0120_NO_NAME_INDEX_OUTPUT",DT_FTR03_0120_NO_NAME_INDEX)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("DocumentNo",-441,-32,"DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("DocumentNo",-364,-32,"DT_FTR03_BLINE_DATE_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_BLINE_DATE_OUTPUT",DT_FTR03_BLINE_DATE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickLabel("DocumentNo", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FTR03_0120_NO_NAME_INDEX,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call SetTextbox("New portfolio","RFIDTRBOE-PORT2","",DT_FTR03_1009_NEW_PORTFOLIO,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot

''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_FTR03_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FTR03_0100_COMPANY_CODE,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FTR03_0100_FISCAL_YEAR,False) 
Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ)
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

'Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ)
'Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call GetTextboxValue("BKPF-BELNR",0,"DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-FTR03----------''''
Call SetTcode(DT_FTR03_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X9","Added to Portfolio",False)
Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_FTR03_1000_DOCUMENT_DATE_OCC1),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_FTR03_1000_POSTING_DATE_OCC1),False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FTR03_1000_COMPANY_CODE_OCC1,False) 
Call TakeScreenShot
Call SelectTab("TABSTRIP_TABS1","Selection Criteria",False)
Call TakeScreenShot

Call SetTextbox("New Portfolio","P_PORT2","",DT_FTR03_1002_NEW_PORTFOLIO,False) 
Call SetTextbox("Old Portfolio","P_PORT4","",DT_FTR03_1002_OLD_PORTFOLIO_OCC1,False) 
Call SetTextbox("Currency","P_WAERS","",DT_FTR03_1002_CURRENCY_OCC1,False) 
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickLabel("DocumentNo", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FTR03_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call GetLabelContentByRefLabel("DocumentNo",-441,-32,"DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",DT_FTR03_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectCheckboxNoLabel(0,DT_FTR03_0120_NO_NAME_OCC1,false)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenshot
Call ClickButton("Yes",True)
Call TakeScreenshot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenshot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot

''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_FTR03_0120_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FTR03_0100_COMPANY_CODE_OCC1,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FTR03_0100_FISCAL_YEAR_OCC1,False) 
Call PressEnter()     
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_UMSKZ_OCC1)
Call VerifyGridCellContent("", 2, "UMSKZ", 0, "")

Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT_OCC1)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)

''Call VerifyGridCellContent("", 2, "UMSKZ", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_UMSKZ_OCC1)
''Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FTR03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT_OCC1)

Call GetTextboxValue("BKPF-BELNR",0,"DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",DT_FTR03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call Logoff()
Call FinalStatus()




