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

'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_072-Physical inventory Non VBS -SAP only RW04_P3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_072-Physical inventory_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_072-Physical inventory  Non VBS -SAP only  RW04_P3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''Login'''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
''''----------------------Tcode MI03----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Phys\. Inventory Doc\.","RM07I-IBLNR","",DT_MI03_0701_PHYS_INVENTORY_DOC,False)
Call SetTextbox("Fiscal Year","RM07I-GJAHR","",DT_MI03_0701_FISCAL_YEAR,False)

Call PressEnter()   
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Physical Inventory History   \(F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Verify details
Call VerifyTextBoxContent("Article","ISEG-MATNR","",DT_MI03_0710_CHECK_TEXT_OF_ARTICLE,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI03_0710_CHECK_TEXT_OF_SITE,False)
Call VerifyTextBoxContent("Storage Loc\.","IKPF-LGORT","",DT_MI03_0710_CHECK_TEXT_OF_STOR_LOCATION,False)
Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",Lcase(DT_MI03_0710_CHECK_TEXT_OF_PHYSINVSTATUS),False)
Call VerifyTextBoxContent("Count date","ISEG-ZLDAT","",DT_MI03_0710_CHECK_TEXT_OF_COUNT_DATE,False)
Call VerifyTextBoxContent("Posting Date","ISEG-BUDAT","",DT_MI03_0710_CHECK_TEXT_OF_POSTING_DATE,False)
Call VerifyTextBoxContent("Quantity","ISEG-MENGE","",DT_MI03_0710_CHECK_TEXT_OF_QUANTITY,False)
Call VerifyTextBoxContent("Book quantity","ISEG-BUCHM","",DT_MI03_0710_CHECK_TEXT_OF_BOOK_QUANTITY,False)
Call VerifyTextBoxContent("Difference qty","VM07I-DIFMG","",DT_MI03_0710_CHECK_TEXT_OF_DIFFERENCE_QTY,False)

Call ClickButton("Next Item   \(Shift\+F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

'Verify details
Call VerifyTextBoxContent("Article","ISEG-MATNR","",DT_MI03_0710_CHECK_TEXT_OF_ARTICLE_OCC1,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI03_0710_CHECK_TEXT_OF_SITE_OCC1,False)
Call VerifyTextBoxContent("Storage Loc\.","IKPF-LGORT","",DT_MI03_0710_CHECK_TEXT_OF_STOR_LOCATION_OCC1,False)
Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",Lcase(DT_MI03_0710_CHECK_TEXT_OF_PHYSINVSTATUS_OCC1),False)
Call VerifyTextBoxContent("Count date","ISEG-ZLDAT","",DT_MI03_0710_CHECK_TEXT_OF_COUNT_DATE_OCC1,False)
Call VerifyTextBoxContent("Posting Date","ISEG-BUDAT","",DT_MI03_0710_CHECK_TEXT_OF_POSTING_DATE_OCC1,False)
Call VerifyTextBoxContent("Quantity","ISEG-MENGE","",DT_MI03_0710_CHECK_TEXT_OF_QUANTITY_OCC1,False)
Call VerifyTextBoxContent("Book quantity","ISEG-BUCHM","",DT_MI03_0710_CHECK_TEXT_OF_BOOK_QUANTITY_OCC1,False)
Call VerifyTextBoxContent("Difference qty","VM07I-DIFMG","",DT_MI03_0710_CHECK_TEXT_OF_DIFFERENCE_QTY_OCC1,False)

Call ClickButton("Display Article Document   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Overview   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Accounting Documents   \(F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("BKPF-BELNR","","DT_MI03_0750_GET_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Verify details
Call VerifyTextBoxContent("Document Date","BKPF-BLDAT","",DT_MI03_0750_CHECK_TEXT_OF_DOCUMENT_DATE,False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_MI03_0750_CHECK_TEXT_OF_COMPANY_CODE,False)
Call VerifyTextBoxContent("Posting Date","BKPF-BUDAT","",DT_MI03_0750_CHECK_TEXT_OF_POSTING_DATE,False)

Call VerifyGridCellContent("",1,"Posting Key","",DT_MI03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Account","",DT_MI03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",1,"Profit Center","",DT_MI03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call VerifyGridCellContent("",2,"Posting Key","",DT_MI03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",2,"Account","",DT_MI03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",2,"Profit Center","",DT_MI03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

