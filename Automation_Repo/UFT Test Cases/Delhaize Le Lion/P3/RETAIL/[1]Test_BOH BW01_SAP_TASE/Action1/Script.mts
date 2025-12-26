
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_BOH BW01_SAP_TASE
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


gstrTestCaseName = "Test_BOH BW01_SAP_TASE"
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''' Login '''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextboxNoLabel("MS_MATNR-LOW","",DT_MMBE_1000_ARTICLE,False)
''Call SetTextbox("to","CRETIM-HIGH","",DT_WE02_1100_CREATED_AT_TO,False)
Call SetTextboxNoLabel("MS_WERKS-LOW","",DT_MMBE_1000_SITE,False)
Call SetTextboxNoLabel("MS_LGORT-LOW","",DT_MMBE_1000_STORAGE_LOCATION,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

'Call ActivateItemGuiTree ("3","","0001 Regular Storage")
Call ActivateItemGuiTree(0,"Full","Full")
Call GetGridContentByRefColumn("Full", 0, "Stock Type", "Unrestricted use", "Stock", "DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE_2_OUTPUT")
Call VerifyGridCellContent("Full",1,"Stock","1",DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE_2_OUTPUT)
Call TakeScreenShot()

Call ClickBUtton("Find   \(Ctrl\+F\)",True)
Call Takescreenshot()
Call SetTextboxNoLabel("GS_SEARCH-VALUE","",DT_SEARCH_TERM,True)
Call SelectCheckbox("GS_SEARCH-EXACT_WORD", 0, DT_ENTIRE_BOX, True)
Call Takescreenshot()
Call ClickBUtton("OK   \(Enter\)",True)
Call Takescreenshot()
Call ClickBUtton("Cancel   \(F12\)",True)
Call Takescreenshot()
Call VerifyGridCellContentByRefColumn("Full", "","Stock Type", DT_SEARCH_TERM, "Stock", 0, DT_MMBE_0300_CHECK_TEXT_OF_TREE_FULL_NODE)
Call ClickButtonifexist("Continue   \(Enter\)",True)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()
