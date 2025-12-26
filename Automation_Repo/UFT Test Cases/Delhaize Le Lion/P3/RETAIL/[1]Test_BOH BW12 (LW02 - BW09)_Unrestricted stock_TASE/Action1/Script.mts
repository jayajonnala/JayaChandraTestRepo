
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_BOH BW12 (LW02 - BW09)_Unrestricted stock_TASE
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


gstrTestCaseName = "TC1_Test_BOH BW12 (LW02 - BW09)_Unrest"
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextboxNoLabel("MS_MATNR-LOW","",DT_MMBE_1000_ARTICLE,False)
Call SetTextboxNoLabel("MS_WERKS-LOW","",DT_MMBE_1000_SITE,False)
Call SetTextboxNoLabel("MS_LGORT-LOW","",DT_MMBE_1000_STORAGE_LOCATION,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call ClickButtonToolBar("DETAILS",0)
Call ClickButton("Find   \(Ctrl\+F\)",True)
Call TakeScreenShot()
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Blocked",True)
wait 2
Call SelectCheckbox("GS_SEARCH-EXACT_WORD",1,"ON",True)
Call TakeScreenShot()
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot()
Call VerifyGridCellContent("Full","1","Stock","",DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE_2)
Call ClickButton("Continue   \(Enter\)",True)

'Call ActivateItemGuiTree ("3","","0001 Regular Storage")
'Call ActivateItemGuiTree(0,"Full","Full")

' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
'Call VerifyGridCellContent("Full",1,"Stock","1",DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE_2)
Call TakeScreenShot()
Call SetTcode(DT_TRANSACTION_CODE2)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Table","GD-TAB","",DT_TABLE_SE16N,false)
Call PressEnter()

Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 2, "", "", DT_DOC_DATE_SE16N, False)
Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 3, "", "", DT_ARTICLE_SE16N, False)
Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 4, "", "", DT_SITE_SE16N, False)
Call TakeScreenShot
Call ClickButton("Online   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell", 1, "Unrestricted", "", DT_CHECK_QTY_WMS)
Call LogOff()
Call FinalStatus ()
