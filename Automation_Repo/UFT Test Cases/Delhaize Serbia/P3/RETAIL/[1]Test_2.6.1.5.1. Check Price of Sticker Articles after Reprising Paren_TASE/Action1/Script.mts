

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_2.6.1.5.1. Check Price of Sticker Articles after Reprising Paren
'.................Author : TCS 	   :
'................ Creation Date    : 
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.6.1.5.1. Check Price of Sticker Articles after Reprising Paren"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.6.1.1.1. Create Article Purchasing Data.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=7
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode VKP5 ----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call ClickButton("Get variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot()

Call SetTextbox("Variant","V-LOW","",DT_VKP2_0100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot()


Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP2_1000_ARTICLE,False)
Call SetTextbox("Validity Period","S_DATUM-LOW","",ConvertDate(DT_VKP2_1000_VALIDITY_PERIOD),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDate(DT_VKP2_1000_TO),False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call VerifyStatusBarMessageType("S")

Call VerifyGridCellContent("Sales Price Conditions", 1, "MATNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("Sales Price Conditions", 1, "VRKME", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VRKME)
Call VerifyGridCellContent("Sales Price Conditions", 1, "VKORG", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VKORG)


Call VerifyGridCellContent("Sales Price Conditions", 1, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 2, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 3, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 4, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 5, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 6, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 7, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 8, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 9, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_8_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 10, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_9_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 11, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_10_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 12, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_11_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 13, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_12_PLTYP)
Call VerifyGridCellContent("Sales Price Conditions", 14, "PLTYP", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_13_PLTYP)


Call VerifyGridCellContent("Sales Price Conditions", 1, "WAERK", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_WAERK)
Call VerifyGridCellContent("Sales Price Conditions", 1, "DATAB", 0, ConvertDate(DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DATAB))

Call VerifyGridCellContent("Sales Price Conditions", 1, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 2, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 3, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 4, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 5, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 6, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 7, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 8, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 9, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 10, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 11, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 12, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 13, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KOTABNR)
Call VerifyGridCellContent("Sales Price Conditions", 14, "KOTABNR", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_13_KOTABNR)



Call VerifyGridCellContent("Sales Price Conditions", 1, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 2, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 3, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 4, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 5, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 6, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 7, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 8, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 9, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 10, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 11, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 12, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 13, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)
Call VerifyGridCellContent("Sales Price Conditions", 14, "KSCHL", 0, DT_VKP2_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KSCHL)


Call LogOff'
Call FinalStatus ()
