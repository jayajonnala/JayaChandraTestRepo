'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-09-Mass assign articles to assortmentsRegularArticle  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-05-09-MRA"
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

''''''--------TransactionCode-ZMDAS_ASSIGN_ASSORT----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Assortment Type","S_ASTYP-LOW","",DT_ZMDAS_ASSIGN_ASSORT_1000_ASSORTMENT_TYPE,False)
Call SetTextbox("Distribution Channel","S_VTWEG-LOW","",DT_ZMDAS_ASSIGN_ASSORT_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Sales Organization","S_VKORG-LOW","",DT_ZMDAS_ASSIGN_ASSORT_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Article","S_MATNR-LOW","",DT_ZMDAS_ASSIGN_ASSORT_1000_ARTICLE,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call ActivateNodeGuiTree(0,"BS10;01;BAOT")
Call TakeScreenShot
Call ClickButtonIfExist("Change/Display   \(F7\)",False)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("",1,1,False)
Call Click204ButtonToolBar("COPY",0)
Call TakeScreenShot
Call ClickButtonIfExist("Complete Data   \(Ctrl\+F3\)",False)
Call Click204ButtonToolBar("&MB_FILTER",0)
Call DoubleClickGuiGridCell("Column Set", "", 5, "Column Name", True)
Call ClickButtonIfExist("Define Filter Values",True)
Call TakeScreenShot
Call SetTextbox("Assortment","%%DYN001-LOW","",DT_ZMDAS_ASSIGN_ASSORT_1105_ASSORTMENT,True)
Call ClickButtonIfExist("Selection Options   \(F2\)",True)
Call DoubleClickGuiGridCell("Assortment", 1, 6, "Selection Options", True)
Call PressEnter()
Call TakeScreenShot
Call SelectAllRowGuiGrid("",1,False)
Call TakeScreenShot
Call Click204ButtonToolBar("DELETE",0)
Call TakeScreenShot
Call GetGridContentByTitle("", "", "Article", 1,DT_OP_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",1,"Assortment",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASORT)
Call VerifyGridCellContent("",1,"Layout Module",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LAYGR)
Call VerifyGridCellContent("",1,"OT",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASSORDIMVAL1)
Call VerifyGridCellContent("",1,"Description",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NAME1)
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
Call ClickButton("Display Log   \(Shift\+F1\)",False)
Call VerifyifGuiLabelExists("Material 00000000000"&DT_ZMDAS_ASSIGN_ASSORT_1000_ARTICLE&" assigned to assortment "&DT_ASSORTMENT)
Call ClickButton("Continue   \(Enter\)", True)
Call TakeScreenShot
Call PressEnter()
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

