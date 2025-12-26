'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-06-01-07-DLE120 vendor conditions report  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-06-01-07-DLE120 vend con"
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

''''''--------TransactionCode-ZMDPC_MARGIN----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButton("Get Variant...   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Variant","V-LOW","",DT_ZMDPC_MARGIN_0600_GRIDCELL_0_VARIANT_NAME,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot

Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDPC_MARGIN_1000_ARTICLE,False)
Call SetTextbox("Valid On Date","P_VKKAB","",ConvertDate(DT_ZMDPC_MARGIN_1000_VALID_ON_DATE),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)

Call VerifyGridCellContent("", 1, "Sales Org", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG)
Call VerifyGridCellContent("", 1, "Dist Chl", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG)
Call VerifyGridCellContent("", 1, "PL", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PLTYP)
Call VerifyGridCellContent("", 1, "Article", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call TakeScreenShot

Call ClickButton("Change Layout...   \(Ctrl\+F8\)",False)
'''Call SelectRowGuiGridbyRowNo("Column Set",0,12, True)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDPC_MARGIN_SEARCH_TERM,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Transfer   \(Enter\)",True)

Call GetGridContent("", "", "Total Margin (%)", 1, "Sales Org", "BS10", "DT_OP_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MARGIN_DLL_T")
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

Call SetTextbox("Valid On Date","P_VKKAB","",ConvertDate(DT_ZMDPC_MARGIN_1000_VALID_ON_DATE_OCC1),False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButton("Change Layout...   \(Ctrl\+F8\)",False)
'''Call SelectRowGuiGridbyRowNo("Column Set",0,1, True)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_ZMDPC_MARGIN_SEARCH_TERM_OCC1,True)
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Transfer   \(Enter\)",True)

Call VerifyGridCellContent("", 1, "Sales Org", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG_OCC1)
Call VerifyGridCellContent("", 1, "Dist Chl", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VTWEG_OCC1)
Call VerifyGridCellContent("", 1, "PL", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PLTYP_OCC1)
Call VerifyGridCellContent("", 1, "Article", "", DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR_OCC1)
Call TakeScreenShot

Call GetGridContent("", "", "Purch Price (PB00)", 1, "Sales Org", "BS10","DT_OP_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BAS_PRICE")
Call GetGridContent("", "", "Net Value (ZPAF)", 1, "Sales Org", "BS10","DT_OP_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NET_PRICE")
Call GetGridContent("", "", "Efective Price (ZDPR)", 1, "Sales Org", "BS10","DT_OP_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EFE_PRICE")
Call GetGridContent("", "", "WS Price (ZPR0) PG", 1, "Sales Org", "BS10","DT_OP_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WHS_PRICE")

Call VerifyGridCellContent("", 1, "Retail Sales UOM", "",DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VRKME)
Call VerifyGridCellContent("", 1, "Total Margin (%)", "",DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MARGIN_DLL_T_OCC1)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

