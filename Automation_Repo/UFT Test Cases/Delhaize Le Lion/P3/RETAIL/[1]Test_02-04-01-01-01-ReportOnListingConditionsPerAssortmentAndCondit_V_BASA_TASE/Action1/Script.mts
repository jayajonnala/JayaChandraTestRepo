'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_02-04-01-01-01-ReportOnListingConditionsPerAssortmentAndCondit_V_BASA  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02-04-01-01-01-ReBASA"
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

''''''--------TransactionCode-WSM4L----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SelectRadioButton("S_DATUM", "Documents for a Date", False)
Call SetTextbox("Assortment","FILIA-LOW","",DT_NEW_ASSORTMENT,False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ActivateItemGuiTree(0, "#2;#1", "#1")
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Message Text","",DT_WSM4L_0101_CHECK_GETCELLVALUE_OF_GRIDCELL_0_T_MSG)
Call ClickButtonIfExist("Cancel   \(F12\)",False)

''''''--------TransactionCode--SM37---------''''

Call SetTcode(DT_WSM4L_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call ClickButtonIfExist("Refresh   \(F8\)",False)
Call TakeScreenShot

''''''--------TransactionCode--/nSE16n---------''''

Call SetTcode(DT_WSM4L_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot


Call SetTextbox("Table","GD-TAB","",DT_WSM4L_0100_TABLE,False)
Call TakeScreenShot
Call PressEnter()
Call SetTableData("SAPLSE16NSELFIELDS_TC", "Fr.Value", 2, "", "", DT_WSM4L_0100_TABLECELL_FRVALUE_1, false)
Call TakeScreenShot
Call PressEnter()
Call ClickButtonIfExist("Online   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Refresh   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContentbyName("shell",1,"Assortment","",DT_WSM4L_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_FILIA)
Call VerifyGridCellContentbyName("shell",1,"Article","",DT_WSM4L_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ARTNR)
Call VerifyTextBoxContent("Search in Table","GD-TAB","",DT_WSM4L_0200_CHECK_TEXT_OF_TABLE_TO_BE_SEARCHED,False)
Call GetGridContentByTitle("", "", "Module", 1, "DT_OP_WSM4L_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_URSAC")
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
