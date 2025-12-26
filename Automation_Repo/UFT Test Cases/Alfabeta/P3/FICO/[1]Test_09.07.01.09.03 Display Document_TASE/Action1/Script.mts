

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.09.03 Display Document
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


gstrTestCaseName = "Test_09.07.01.09.03 Display Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''''''--------TransactionCode-FB03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year", "RF05L-GJAHR", "", Year(DT_FB03_0100_FISCAL_YEAR), False)
Call PressEnter()     
Call TakeScreenShot

Call VerifyTextBoxContent("Document Number", "BKPF-BELNR", "", DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER, False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS", "", DT_FB03_0750_CHECK_TEXT_OF_COMPANY_CODE, False)
Call VerifyTextBoxContent("Document Date", "BKPF-BLDAT", "", ConvertDate(DT_FB03_0750_CHECK_TEXT_OF_DOCUMENT_DATE), False)
Call VerifyTextBoxContent("Posting Date", "BKPF-BUDAT", "", ConvertDate(DT_FB03_0750_CHECK_TEXT_OF_POSTING_DATE), False)

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


Call VerifyGridCellContent("", 1, "G/L Account", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "LOKKT", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("", 1, "KOBEZ", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOBEZ)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("", 1, "Currency", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("", 1, "PRCTR", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 1, "Text", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call VerifyGridCellContent("", 2, "G/L Account", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "LOKKT", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("", 2, "KOBEZ", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOBEZ)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("", 2, "Currency", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_RF05A_UBAZW)
Call VerifyGridCellContent("", 2, "MWSKZ", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 2, "GSBER", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)
Call VerifyGridCellContent("", 2, "PRCTR", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 2, "Functional Area", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_FKBER_LONG)
Call VerifyGridCellContent("", 2, "Text", 0, DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call LogOff'
Call FinalStatus()


