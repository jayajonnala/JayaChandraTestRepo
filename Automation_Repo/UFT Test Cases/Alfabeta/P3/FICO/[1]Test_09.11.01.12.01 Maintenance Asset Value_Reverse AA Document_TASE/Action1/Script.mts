
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'''''''.................Test Script Name :Test_09.11.01.12.01 Maintenance Asset Value_Reverse AA Document
'''''''.................Author : TCS 
'''''''................ Creation Date :
'''''''.................Modified By :
'''''''.................Modified Date/Details :
'''''''
'''''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
''''''

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

gstrTestCaseName = "Test_09.11.01.12.01 Maintenance Asset Value_Reverse AA Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


''''''''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'''''DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''--------------login----------------'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''--------TransactionCode-AB08----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("Company Code","RLAB01-BUKRS","",DT_AB08_0010_COMPANY_CODE,False)
Call SetTextbox("Asset","RLAB01-ANLN1","",(DT_AB08_0010_ASSET),False)
Call SetTextbox("Sub-number","RLAB01-ANLN2","",DT_AB08_0010_SUBNUMBER,False)
Call SetTextbox("Fiscal Year","RLAB01-GJAHR","",DT_AB08_0010_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Reverse Document   \(F6\)",False)
Call TakeScreenShot
Call SetTextbox("Document Number","RF05A-BELNS","",DT_AB08_0105_DOCUMENT_NUMBER,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",(DT_AB08_0105_COMPANY_CODE),False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_AB08_0105_FISCAL_YEAR,False)
Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_AB08_0105_REVERSAL_REASON,False)
Call SetTextbox("Posting Date","BSIS-BUDAT","",ConvertDate(DT_AB08_0105_POSTING_DATE),False)
Call SetTextbox("Posting period","BSIS-MONAT","",DT_AB08_0105_POSTING_PERIOD,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Display document before reversal   \(F5\)",False)
Call TakeScreenShot

Call Verifytextboxcontent("Document Number","BKPF-BELNR",0,DT_AB08_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_DOCUMENT_AB08_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_DOCUMENT_AB08_OUTPUT",DT_DOCUMENT_AB08)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call VerifyStatusBar(Lcase(DT_SUCCESSMESSAGE_AB08))

''''--------TransactionCode-AS03----------''''
Call SetTcode(DT_AB08_0010_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#1;#3")

Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)

Call ActivateNodeGuiTree(0, "#1;#1")

Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT))
Call VerifyGridCellContent("Transactions", 2, "BWATXT", 0, DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWATXT)

Call DoubleClickGuiGridCell("Transactions", 0, 2, "BWATXT", False)

Call VerifyGridCellContent("", 1, "Negative posting", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_XNEGP)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "Account", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "Amount", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyGridCellContent("", 2, "Negative posting", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_XNEGP)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "Account", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 3, "Negative posting", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_XNEGP)
Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 3, "Account", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)

''''--------TransactionCode-FBL1N----------''''
Call SetTcode(DT_AB08_0750_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_AB08_1000_VENDOR_ACCOUNT,False)
Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_AB08_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call SelectRadioButton("X_AISEL","All items",False)
Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)

Call ClickButtonIfExist("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_AB08_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_AB08_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)

Call LogOff'
Call FinalStatus()
