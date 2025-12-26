'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.10.01 Retire Assets_Asset Retirement by Scrapping_P2
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

gstrTestCaseName = "Test_09.11.01.10.01 Retire Assets_Asset Retirement by Scrapping_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-AS01 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All Items",False)
Call TakeScreenShot()

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False) 
Wait(2)
Call TakeScreenShot()
Wait(2)

'
''Click on GUI Tree
'''Call ActivateNodeGuiTree(0,"General Ledger Line Items;Document Number")
Call ActivateNodeGuiTree(0,"G/L Account Line Items;Document Number")
Wait(2)

Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLL03_0100_DOCUMENT_NUMBER,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Yes",True)
wait(2)

Call ClickButton("Choose Ledger   \(Ctrl\+F3\)",False) 
Wait(2)

Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLL03_0300_LEDGER,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True) 
Wait(2)

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Wait(2)
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Verify the grid Content
Call VerifyGridCellContent("",1,"Document Number",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Document type",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("",1,"Document Date",0,ConvertDate(DT_DATE_REF))
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",1,"Text",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButton("Choose Ledger   \(Ctrl\+F3\)",False) 
Wait(2)

Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLL03_0300_LEDGER_OCC1,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True) 
Wait(2)


'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Verify the grid Content
Call VerifyGridCellContent("",1,"Document Number",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Document type",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("",1,"Document Date",0,ConvertDate(DT_DATE_REF))
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",1,"Text",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT_OCC1,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE_OCC1,False)

Call ClickButton("Choose Ledger   \(Ctrl\+F3\)",False) 
Wait(2)

Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLL03_0300_LEDGER_OCC1,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True) 
Wait(2)

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()


'Verify the grid Content
Call VerifyGridCellContent("",1,"Document type",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("",1,"Document Date",0,ConvertDate(DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1))
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("",1,"Text",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


Call ClickButton("Choose Ledger   \(Ctrl\+F3\)",False) 
Wait(2)

Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLL03_0300_LEDGER,True)
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True) 
Wait(2)

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()


'Verify the grid Content
Call VerifyGridCellContent("",1,"Document type",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("",1,"Document Date",0,ConvertDate(DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1))
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("",1,"Text",0,DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


''----------------------Tcode AS03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call SetTcode(DT_FAGLL03_1000_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_FAGLL03_1000_OKCD)
Call PressEnter()

'Fill The details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_FAGLL03_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_FAGLL03_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_FAGLL03_0100_COMPANY_CODE,False)
Call TakeScreenShot

'Click on Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)


Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_FAGLL03_0304_CHECK_TEXT_OF_USEFUL_LIFE,False)

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Planned values",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)


Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_FAGLL03_0304_CHECK_TEXT_OF_USEFUL_LIFE_OCC1,False)

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Planned values",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;L1 Ledger (Local IFRS Gap);22 IFRS LOCL GAAP APC, depreciation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)


Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_CHECK_22_2,False)

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Planned values",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;20 Pre-merger STA valuation")

'Navigate to Parameters Tab
Call SelectTab("IDC_TABSTRIP","Parameters",False)

Call VerifyTextBoxContent("Useful life","AW01_DEP_PAR-NDJAR",0,DT_CHECK_20_2,False)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

