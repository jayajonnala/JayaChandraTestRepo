'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-05-09-Mass assign articles to assortments-Reg_BE_V_BASA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-05-09-MaBEBASA"
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
Call ActivateNodeGuiTree("0","BS10;01;BASA")
Call TakeScreenShot
Call ClickButtonIfExist("Change/Display   \(F7\)",False)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("",1,1,False)
Call Click204ButtonToolBar("COPY",0)
Call TakeScreenShot
Call SelectColumnGuiGrid("",1,"SA",False)
Call Click204ButtonToolBar("&MB_FILTER",0)
Call ClickButtonIfExist("\%\_\%\%DYN001\_\%\_APP\_\%\-VALU\_PUSH",True)
Call ClickButtonIfExist("Define Selection Options   \(F2\)",True)
Call TakeScreenShot

Call ClickButtonIfExist("Copy   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call SetGridData("",1,"SA",DT_ZMDAS_ASSIGN_ASSORT_0100_GRIDCELL_0_SA,False)
Call TakeScreenShot
Call ClickButtonIfExist("Complete Data   \(Ctrl\+F3\)",False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
'Call VerifyTextPopup(DT_ZMDAS_ASSIGN_ASSORT_0010_CHECK_TEXT_OF_MESSTXT1) ' Pop up screen is not coming.
Call TakeScreenShot
'Call ClickButton("Continue \(Enter\)", True)  'Step will be enabled when Pop-up screen will come.
Call ClickButton("Display Log   \(Shift\+F1\)",False)
Call TakeScreenShot
Call VerifyifGuiLabelExists("Material 000000000007167528 assigned to assortment "&DT_NEW_ASSORTMENT)
Call ClickButton("Continue   \(Enter\)", True)
Call VerifyGridCellContent("",2,"Article",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",2,"SA",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASSORDIMVAL1)
Call VerifyGridCellContent("",2,"Description",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASORT)
Call VerifyGridCellContent("",2,"Layout Module",1,DT_ZMDAS_ASSIGN_ASSORT_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ASORT)
Call TakeScreenShot
Call PressEnter()
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

