'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_09.Credit note - no GI  for an affiliated store in LU54  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.Credit note - no GI   LU54"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''--------TransactionCode-va03----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE) 

Call SetTextbox("Order","VBAK-VBELN","",DT_VA03_0102_ORDER,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Display document flow   \(F5\)",False)
Call SelectMenuBar("Document flow;View;Items")
Call ActivateItemGuiTree(0, "#2;#1;#1", "#1")
Call GetGridContent("Accounting document.*", 0, "Doc.no.", 1, "Status", "Not Cleared", "DT_OP_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM")
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)

'''''''--------TransactionCode-/FB03----------''''


Call SetTcode(DT_VA03_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_VA03_0100_FISCAL_YEAR,False)
Call SetTextbox("Document Number","RF05L-BELNR","",DT_VA03_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_VA03_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyGridCellContent("", 1, "Posting Key","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "Posting Key","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 3, "Posting Key","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 1, "Account","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 2, "Account","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("", 3, "Account","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("", 1, "Tax Code","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
Call VerifyGridCellContent("", 2, "Tax Code","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 3, "Tax Code","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call ClickButtonToolBar("&MB_FILTER",1)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Column Set", 0, 36, True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND", True)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Trading Partner",True)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0, "ON", True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Column Set", 0, 74, True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("Column Set", 0,22, True)
Call ClickButtonIfExist("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Transfer   \(Enter\)",True)
Call VerifyGridCellContent("", 1, "Business Area","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
Call VerifyGridCellContent("", 2, "Business Area","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)
Call VerifyGridCellContent("", 3, "Business Area","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_GSBER)
Call VerifyGridCellContent("", 1, "Assignment","","")
'Call VerifyGridCellContent("", 2, "Assignment","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)
'Call VerifyGridCellContent("", 3, "Assignment","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ZUONR)
'Call VerifyGridCellContent("", 1, "Cost Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KOSTL)
'Call VerifyGridCellContent("", 2, "Cost Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KOSTL)
'Call VerifyGridCellContent("", 3, "Cost Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KOSTL)
Call VerifyGridCellContent("", 1, "Profit Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("", 2, "Profit Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
Call VerifyGridCellContent("", 3, "Profit Center","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PRCTR)
Call VerifyGridCellContent("", 1, "Alternative Account No.","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 2, "Alternative Account No.","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("", 3, "Alternative Account No.","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)
'Call VerifyGridCellContent("", 1, "Trading partner","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBUND)
'Call VerifyGridCellContent("", 2, "Trading partner","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_VBUND)
'Call VerifyGridCellContent("", 3, "Trading partner","",DT_VA03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_VBUND)
Call ClickButtonIfExist("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type", "BKPF-BLART", "", lcase(DT_VA03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE), True)
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

