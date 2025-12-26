'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.10.01 Retire Assets_Asset Retirement by Scrapping [Unique]_P1
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

gstrTestCaseName = "Test_09.11.01.10.01 Retire Assets_Asset Retirement by Scrapping [Unique]_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-ABAVN ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter Company Code
Call SetTextbox("Company Code","SVALD-VALUE","",DT_ABAVN_0300_COMPANY_CODE,True)
'Click on Enter
Call ClickButton("Continue   \(Enter\)",True) 
Wait(2)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Asset","RAIFP2-ANLN1","",DT_ABAVN_0300_ASSET,False)
Call SetTextbox("Document Date","RAIFP1-BLDAT","",ConvertDate(DT_ABAVN_0200_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","RAIFP1-BUDAT","",ConvertDate(DT_ABAVN_0200_DOCUMENT_DATE),False)
Call SetTextbox("Asset Value Date","RAIFP1-BZDAT","",ConvertDate(DT_ABAVN_0202_ASSET_VALUE_DATE),False)
Call SetTextbox("Text","RAIFP2-SGTXT","",DT_ABAVN_0206_TEXT,False)
Call TakeScreenShot
Call PressEnter() 

'Call GetTextboxValue("RAIFP2-SGTXT",0,"DT_TEXT_OUTPUT",False)
Call SelectTab("TABSTRIP100", "Partial retirement", False)


Call SetTextbox("Amount posted","RAIFP2-ANBTR","",DT_ABAVN_0401_AMOUNT_POSTED,False)
Call SelectRadioButton("RAIFP2-XANEU","From curr\.-yr aquis\.",False)
Call TakeScreenShot


Call ClickButton("Simulate   \(F9\)",False)

Call VerifyGridCellContent("Line items",1,"Amount",0,DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WRBTR)
Call VerifyGridCellContent("Line items",2,"Amount",0,DT_ABAVN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_WRBTR)


'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item2","DT_ASSET_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSET_OUTPUT",DT_ASSET)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
''Call VerifyStatusBar("Asset transaction posted with document no. GR02 "&DT_ASSET_TRANSACTION_NO_OUTPUT)
Call VerifyStatusBar("Asset transaction posted with document no. GR02 "&DT_ASSET_OUTPUT)

''----------------------Tcode AS03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call SetTcode(DT_ABAVN_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_ABAVN_0100_OKCD)
'Call PressEnter()

'Fill The details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABAVN_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABAVN_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABAVN_0100_COMPANY_CODE,False)
Call TakeScreenShot

'Click on Master Data
Call ClickButton("Master data   \(F7\)",False)
Wait(3)
Call TakeScreenShot


'Click on Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
wait(3)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)

Call VerifyGridCellContent("Transactions",2,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT))
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)
Call VerifyGridCellContent("Transactions",2,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL)

Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")

Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC1)

Call VerifyGridCellContent("Transactions",2,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions",2,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC1)

Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")


Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC1)

Call VerifyGridCellContent("Transactions",2,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions",2,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC1)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;20 Pre-merger STA valuation")

Call VerifyGridCellContent("Transactions",1,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions",1,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC1)

Call VerifyGridCellContent("Transactions",2,"Asset Value Date",0,ConvertDate(DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC1))
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC1)
Call VerifyGridCellContent("Transactions",2,"Transaction Type",0,DT_ABAVN_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC1)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

Call DoubleClickGuiGridCell("Transactions",0,2,"Transaction Type",False)

Call VerifyGridCellContent("",1,"Account",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("",2,"Account",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("",1,"Currency",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PSWSL)
Call VerifyGridCellContent("",2,"Currency",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)

Call VerifyGridCellContent("",1,"Amount",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"Amount",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("",1,"Posting Key",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_ABAVN_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call GetGridContent("",0,"Text",1,"Posting Key",75,"DT_ASSET_TEXT_OUTPUT")

Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
wait(2)
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
wait(2)
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
wait(2)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

