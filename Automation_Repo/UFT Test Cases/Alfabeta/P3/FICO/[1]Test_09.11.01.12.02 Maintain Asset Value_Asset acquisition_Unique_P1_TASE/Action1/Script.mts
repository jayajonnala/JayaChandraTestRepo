

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P1
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

gstrTestCaseName = "Test_09.11.01.12.02 Maintain Asset Value_Asset acquisition_Unique_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'----------------------Login----------------------------
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''----------------------Tcode ABSO----------------------------
'
''Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","ANBZ-BUKRS","",DT_ABSO_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANBZ-ANLN1","",DT_ABSO_0100_ASSET,False)
Call SetTextbox("Sub-number","ANBZ-ANLN2","",DT_ABSO_0100_SUBNUMBER,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_ABSO_0100_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_ABSO_0100_POSTING_DATE),False)
Call SetTextbox("Posting period","ANBZ-PERID","",DT_ABSO_0100_POSTING_PERIOD,False)
Call SetTextbox("Transaction Type","ANBZ-BWASL","",DT_ABSO_0100_TRANSACTION_TYPE,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()


Call SetTextbox("Quantity","ANBZ-MENGE","",DT_ABSO_0110_QUANTITY,False)
Call SetTextbox("Amount posted","ANBZ-DMBTR","",DT_ABSO_0110_AMOUNT_POSTED,False)
Call SetTextbox("Text","ANEK-SGTXT","",DT_ABSO_0110_TEXT,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()

Call GetTextboxValue("ANEK-SGTXT", "", "DT_ABSO_0110_CHECK_TEXT_OF_TEXT_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ABSO_0110_CHECK_TEXT_OF_TEXT_OUTPUT",DT_ABSO_0110_CHECK_TEXT_OF_TEXT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot()
Call ClickButton("Line Items   \(Shift\+F1\)",False)
Call TakeScreenShot()

Call VerifyTableCellContent(1, "Area", "SAPMA01BTCTRL_ANBTR", DT_ABSO_0250_CHECK_TEXT_OF_TABLECELL_AREA_0)
Call VerifyTableCellContent(1, "Tra", "SAPMA01BTCTRL_ANBTR", DT_ABSO_0250_CHECK_TEXT_OF_TABLECELL_TRA_0)
Call VerifyTableCellContent(1, "Amount posted", "SAPMA01BTCTRL_ANBTR", DT_ABSO_0250_CHECK_TEXT_OF_TABLECELL_AMOUNT_POSTED_0)

Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot()
Call SetTextbox("Quantity","ANBZ-MENGE","","",False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_ABSO_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction was posted with AA document number "&DT_ABSO_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call WriteRunTimeDataToExcelGlobalSheet("DT_ABSO_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_ABSO_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''''----------------------Tcode AS03----------------------------
Call SetTcode(DT_ABSO_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot    

Call SetTextbox("Asset","ANLA-ANLN1","",DT_ABSO_0100_ASSET_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_ABSO_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_ABSO_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot   

Call ClickBUtton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot 

Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "VERAENDERUNG", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 1, "JENDE", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call ActivateNodeGuiTree(0, "#1;#2;#1")

Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "VERAENDERUNG", 0, "")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 1, "JENDE", 0, "")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "VERAENDERUNG", 0, "")
Call VerifyGridCellContent("Planned values Local GAAP APC, depreciation", 4, "JENDE", 0, "")


Call ActivateNodeGuiTree(0, "#1;#1;#2")

Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "VERAENDERUNG", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 1, "JENDE", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "VERAENDERUNG", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values Pre-merger STA valuation", 4, "JENDE", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call ActivateNodeGuiTree(0, "#1;#3;#1")

Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "VERAENDERUNG", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 1, "JENDE", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "VERAENDERUNG", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS LOCL GAAP APC, depreciation", 4, "JENDE", 0, DT_ABSO_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call ClickButton("Exit   \(Shift\+F3\)",False)

Call ClickButton("Exit   \(Shift\+F3\)",False)


''------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************




