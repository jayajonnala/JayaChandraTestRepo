
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'----------------------------------------------------------------------------------------------------------------------------
gstrTestCaseName = "TC3_Test_02-06-01-09-DLR011 marrep"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Login 
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------VL10G----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Sales Organization","S_VKORG-LOW","",DT_ZMDPC_MARGIN_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution Channel","S_VTWEG-LOW","",DT_ZMDPC_MARGIN_1000_DISTRIBUTION_CHANNEL,False)
Call TakeScreenShot
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDPC_MARGIN_1000_ARTICLE,False)
Call SetTextbox("Valid On Date","P_VKKAB","",DT_ZMDPC_MARGIN_1000_VALID_ON_DATE,False)
Call TakeScreenShot
Call SetTextbox("Price List","S_PLTYP-LOW","",DT_ZMDPC_MARGIN_1000_PRICE_LIST,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)
Call TakeScreenShot
''' SelectCellGuiGrid(gridTitle, gridIndex, Rowno, columnName, blnIsItPopup)
Call SelectCellGuiGrid("Column Set","",0,"Column Name",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenShot
Call SelectCellGuiGrid("Column Set","",5,"Column Name",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call SelectCellGuiGrid("Column Set","",6,"Column Name",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenShot
Call SelectCellGuiGrid("Column Set","",4,"Column Name",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Transfer   \(Enter\)",True)
Call GetGridContentByRefColumn("","","Sales Org","BS10","Efective Price (ZDPR)","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EFE_PRICE_Output")
Call GetGridContentByRefColumn("","","Sales Org","BS10","DLL Margin (Amt)","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MARKUP_DLL_Output")
Call TakeScreenShot
Call GetGridContentByRefColumn("","","Sales Org","BS10","Affiliated Margin (Amt)","DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MARKUP_AFL_Output")
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell",1,"WS Price (ZPR0) PG","",DT_WHS_PRICE)
Call VerifyGridCellContentbyName("shell",1,"Retail Price without VAT (PV HTVA)","",DT_RET_PRICE)
Call VerifyGridCellContentbyName("shell",1,"Previous Retail Price without VAT","",DT_ZMDPC_MARGIN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRR_PRICE)

Call LogOff()
Call FinalStatus()

