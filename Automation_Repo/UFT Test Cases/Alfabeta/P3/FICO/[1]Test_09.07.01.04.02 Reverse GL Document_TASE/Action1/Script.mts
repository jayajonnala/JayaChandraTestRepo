		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.04.02 Reverse GL Document
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


gstrTestCaseName = "Test_09.07.01.04.02 Reverse GL Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''''--------TransactionCode-FB08----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call SetTextbox("Posting Date","BSIS-BUDAT", "", ConvertDate(DT_FB08_0105_POSTING_DATE), False)
Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",Year(DT_FB08_0105_FISCAL_YEAR),False)
Call SetTextbox("Reversal Reason","UF05A-STGRD", "", DT_FB08_0105_REVERSAL_REASON, False)
Call SetTextbox("Posting period","BSIS-MONAT", "", ConvertDoubleDigit(Cstr(MOnth(DT_FB08_0105_POSTING_PERIOD))), False)

Call ClickButton("Display document before reversal   \(F5\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Document Number","BKPF-BELNR", "", DT_FB08_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call ClickButton("Back   \(F3\)",False)
Wait 2
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


''''''''--------TransactionCode-FAGLL03----------''''

Call SetTcode(DT_FB08_0105_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL", "All Items", False)

Call ClickButton("%_SD_SAKNR_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB08_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB08_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#4;#1")

Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB08_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB08_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2

Call SetTextbox("Layout","PA_VARI", "",DT_LAYOUT, False)

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 5, "BELNR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BELNR)
Call VerifyGridCellContent("", 6, "BELNR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BELNR)

Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 5, "BSCHL", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("", 6, "BSCHL", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)

Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DMSHB)
Call VerifyGridCellContent("", 6, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
Call VerifyGridCellContent("", 7, "DMSHB", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_DMSHB)

Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 5, "PRCTR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PRCTR)
Call VerifyGridCellContent("", 6, "PRCTR", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_PRCTR)

Call VerifyGridCellContent("", 4, "ICO_AUGP", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ICO_AUGP)
Call VerifyGridCellContent("", 8, "ICO_AUGP", 0, DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL73_ICO_AUGP)

Call Logoff
Call finalstatus
