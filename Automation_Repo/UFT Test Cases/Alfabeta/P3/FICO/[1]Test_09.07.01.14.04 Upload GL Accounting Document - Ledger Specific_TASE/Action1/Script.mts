

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.14.04 Upload GL Accounting Document - Ledger Specific
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

gstrTestCaseName = "Test_09.07.01.14.04 Upload GL Accounting Document - Ledger Specific"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False) 
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Yes",True)
Call TakeScreenShot()

''''''--------TransactionCode-SM35----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call FindRowNumber("SAPMSBDC_CCTC_APQI", "Created By", "AUTOFAB", "DT_ROW_OUTPUT")

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", DT_ROW_OUTPUT, False)
Call TakeScreenShot
Call ClickButton("Process session   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButtonifExist("Process   \(Enter\)",True)
Call TakeScreenShot

Call SetTextbox("Document Date","BKPF-BLDAT", 0, ConvertDate(DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_DATE), False)
Call SetTextbox("Posting Date","BKPF-BUDAT", 0, ConvertDate(DT_ZFIGL_UPLOAD_POST_0100_POSTING_DATE), False)
'''Call SetTextbox("Period","BKPF-MONAT", 0, ConvertDoubleDigit((Month(DT_ZFIGL_UPLOAD_POST_0100_PERIOD))), False)
Call SetTextbox("Period","BKPF-MONAT", 0, DT_FORMULA_MONTH, False)
Call SetTextbox("Ledger Grp","BKPF-LDGRP", 0, (DT_ZFIGL_UPLOAD_POST_0100_LEDGER_GRP), False)
Call SetTextbox("Translation dte","BKPF-WWERT", 0, ConvertDate(DT_ZFIGL_UPLOAD_POST_0100_TRANSLATN_DATE), False)
Call TakeScreenShot
Call PressEnter()
Call takeScreenShot()
Call SetTextbox("PstKy","RF05A-NEWBS", 0, (DT_ZFIGL_UPLOAD_POST_0300_PSTKY), False)
Call SetTextbox("Account","RF05A-NEWKO", 0, (DT_ZFIGL_UPLOAD_POST_0300_ACCOUNT), False)
Call TakeScreenShot
Call PressEnter()
Call takeScreenShot()
Call PressEnter()
Call takeScreenShot()
Call PressEnter()
Call takeScreenShot()
Call PressEnter()
Call takeScreenShot()

Call PressEnter()
Call takeScreenShot()

Call ClickButtonifExist("Go back to batch input session overview   \(Enter\)",True)
Call TakeScreenShot

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", DT_ROW_OUTPUT, False)
Call TakeScreenShot
Call ClickButton("Analyze session   \(F2\)",False)
Call SelectTab("TAB_DYNPRO", " Dynpros", False)
Call TakeScreenShot'
DT_TABNAME = " Log created on "&ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM)
Call SelectTab("TAB_DYNPRO", DT_TABNAME, False)
Call TakeScreenShot' 

Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL", "Message","13", "", "", "DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_12_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_12_OUTPUT",DT_ZFIGL_UPLOAD_POST_0400_CHECK_TEXT_OF_TABLECELL_MESSAGE_12)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


'''''--------TransactionCode-FB03 ----------''''
Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIGL_UPLOAD_POST_0100_DOCUMENT_NUMBER,False) 
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIGL_UPLOAD_POST_0100_COMPANY_CODE,False) 
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",Year(DT_ZFIGL_UPLOAD_POST_0100_FISCAL_YEAR),False) 
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Display Document Header   \(F5\)",False)

Call VerifyTextBoxContent("Document type","BKPF-BLART", 0, DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call VerifyTextBoxContent("Ledger Grp","BKPF-LDGRP", 0, DT_ZFIGL_UPLOAD_POST_1710_CHECK_TEXT_OF_LEDGER_GRP, True)
Call ClickButton("Continue/Confirm   \(Enter\)",True)

Call SelectMenuIdToolBar("&COL0",1)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","G/L Account",True)
Call SetComboByKey("Search Direction",0)
Call TakeScreenShot
Call PressEnter
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Call ClickButtonIfExist("Show Selected Fields \(F7\)",True)
Call ClickButtonIfExist("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 3, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
Call VerifyGridCellContent("", 4, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_HKONT)
Call VerifyGridCellContent("", 5, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_HKONT)
Call VerifyGridCellContent("", 6, "G/L Account", 0, DT_ZFIGL_UPLOAD_POST_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_HKONT)

Call LogOff()
Call FinalStatus()

