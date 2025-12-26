'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'.................Test Script Name : Test_P2P_01_01_091-Repricing for returns - price_P3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 14th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_091_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_091-Repricing for returns - price_P3_TASE.xls"

'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

'Enter the PO Number and Press Enter
'Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
'Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_SELECT_PURCH_ORDER,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",Replace((DT_MIGO_0110_DOCUMENT_DATE),"/","."),False) 
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",Replace((DT_MIGO_0110_POSTING_DATE),"/","."),False) 
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MIGO_0304_ITEM_OK,False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",1,DT_MIGO_0304_ITEM_OK,False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",2,DT_MIGO_0200_TABLECELL_OK_1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC2)
'_______________________________________________________________________________________________________

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC1)
'Capture the screenshot
Call TakeScreenShot()

'Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC1)
''Capture the screenshot
'Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_2010_GODYNPRODOC_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

Call VerifyTableCellContent(1,"Profit Center","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_EUN_0)
Call VerifyTableCellContent(2,"Profit Center","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_EUN_1)

Call VerifyTableCellContent(1,"Article","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_0)
Call VerifyTableCellContent(2,"Article","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_1)

Call VerifyTableCellContent(1,"Qty in UnE","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_OK_0)
Call VerifyTableCellContent(2,"Qty in UnE","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_OK_1)

Call VerifyTableCellContent(1,"Art. Short Text","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ART_SHORT_TEXT_0)
Call VerifyTableCellContent(2,"Art. Short Text","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ART_SHORT_TEXT_1)

'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Where",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_VENDOR_BATCH_0,False)
Call ClickButton("Next Item",False) 
wait(2)
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_VENDOR_BATCH_1,False)
Call ClickButton("Next Item",False) 

Call SelectTab("TS_GOITEM","Partner",False)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Vendor","GOITEM-LIFNR","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_CUSTOMER_1,False)
Call ClickButton("Previous Item",False) 
wait(2)
Call VerifyTextBoxContent("Vendor","GOITEM-LIFNR","",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_CUSTOMER_0,False)

'
''''----------------------Tcode ME23N----------------------------
'Enter the Tcode
Call SetTcode(DT_ME23N_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_OKCD)
'Capture the screenshot
Call TakeScreenShot()

'Enter the Details
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
'Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_SELECT_PURCH_ORDER,True)
Call TakeScreenShot()
Call ClickButton("Other Document   \(Enter\)",True)

'close document overview if exist
Call ClickButtonIfExist("Switch Off Document Overview   \(F9\)",False)

'Close help page if exist
If SAPGuiSession("transaction:="&DT_SAPTRANSACTIONCODE).SAPGuiWindow("transaction:="&DT_SAPTRANSACTIONCODE).Page("title:=").Exist = True Then	
	Call ClickButton("Help   \(Shift\+F1\)",False)
	Wait(1)	
End If

'navigate to item details section
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)

'Purchase order history tab
Call SelectTab("ITEM_DETAIL",DT_ME23N_1303_PURCHASE_ORDER_HISTORY,False)
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Article Document","",DT_MIGO_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Posting Date","",ConvertDate(DT_MIGO_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT))

'get data
Call GetTableCellData("SAPLMEGUITC_1211","Net Price",1,"Article",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_SITE_0,"DT_MIGO_1211_GET_TEXT_OF_TABLECELL_NET_PRICE_0_OUTPUT",False)
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
'VerifyGridCellContent
' VerifyGridCellContentByRefColumn(gridTitle, gridRowNumber, refColumnName, refCellValue, gridColumnName, gridIndex, expectedValue)
call VerifyGridCellContentByRefColumn("","","Delivery cost quantity","0","Amount",0,DT_MIGO_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)


'Call VerifyGridCellContent("",1,"Amount","",DT_MIGO_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMBTR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


