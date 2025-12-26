'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-03-06-MassDelistingOfArticle_fromAssortment-Reg_BE_BASA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-03-06-MaBASA"
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
Call ActivateNodeGuiTree("","BS10;01;BASA")
Call TakeScreenShot
Call ClickButtonIfExist("Change/Display   \(F7\)",False)
Call TakeScreenShot
Call SelectAllRowGuiGrid("",1,False)
Call Click204ButtonToolBar("&MB_FILTER",0)
Call SetTextbox("Assortment","%%DYN005-LOW","",DT_ZMDAS_ASSIGN_ASSORT_1105_ASSORTMENT,True)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(Enter\)",True)
Call TakeScreenShot
Call SelectRowGuiGridbyRowNo("", 1, 1, False)
Call Click204ButtonToolBar("DELETE",0)
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call TakeScreenShot
'Call VerifyTextPopup(DT_ZMDAS_ASSIGN_ASSORT_0010_CHECK_TEXT_OF_MESSTXT1) ' Pop up screen is not coming.
Call TakeScreenShot
'Call PressEnter() ' Pop up screen is not coming.
Call ClickButton("Display Log   \(Shift\+F1\)",False)
Call VerifyifGuiLabelExists("Material 000000000007167528 deleted from assortment "&DT_NEW_ASSORTMENT)
Call TakeScreenShot
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

