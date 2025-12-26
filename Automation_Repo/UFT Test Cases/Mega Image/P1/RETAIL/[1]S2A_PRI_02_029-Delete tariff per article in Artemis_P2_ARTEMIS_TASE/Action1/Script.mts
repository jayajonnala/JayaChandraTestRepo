
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
'.................Test Script Name :S2A_PRI_02_029-Delete tariff per article in Artemis_P2_ARTEMIS_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "S2A_PRI_02_029-Delete _P2_ARTEMIS_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_S2A_PRI_02_023-Maintain margin driven retail price for tariff 49 ShopandGo_Artemis.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''''--------------login----------------'''''

Call Launch_Mainframe(DT_ENVIRONMENT,"A","PC5250","&Yes")
Call Verify_Launch_LoginPage("A")
Call Login_Mainframe("A",DT_USERNAME,DT_PASSWORD)
Call SendTEScreenKey ("A",DT_P_KEY)
Call Verify_TEFieldExists("MEGA IMAGE SA","","A","","PC5250","&No","LogIN Page")
Call SendTEScreenKey ("A",DT_P_KEY)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME,DT_P_FIELD_ID_VAL_ALEGE)
Call SendTEScreenKey ("A",DT_P_SEND_KEY)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C1,DT_P_FIELD_ID_VAL_ALEGE_C1)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C1)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C2,DT_P_FIELD_ID_VAL_ALEGE_C2)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C2)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C3,DT_P_FIELD_ID_VAL_ALEGE_C3)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C3)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID("A",DT_P_FIELD_ID, DT_P_FIELD_ID_VAL_DAY_C4,"")
Call SetTEFieldBY_ID("A",DT_P_FIELD_ID_C1, DT_P_FIELD_ID_VAL_MONTH_C5,"")
'Call SetTEFieldBY_ID("A",DT_P_FIELD_ID_C2, DT_P_FIELD_ID_VAL_YEAR_C6,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C4)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C4,DT_P_FIELD_ID_VAL_MASTER_ITEM_C7)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C5)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID("A",DT_P_FIELD_ID_C3, DT_P_FIELD_ID_VAL_TIP_PRET_C8,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C6)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C5,DT_P_FIELD_ID_VAL_TARIF_C9)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C7)
Call SendTEScreenKey ("A",DT_P_KEY_C1)
Call CaptureScreenshot_Mainframe("A")
Call VerifyTEFieldByID("A",DT_P_FIELD_ID_C4,DT_P_FIELD_ID_VERIFY_TARIF_C10,"")
Call VerifyTEFieldByID("A",DT_P_FIELD_ID_C5,DT_P_FIELD_ID_VERIFY_PRICE_TYPE_C11,"")
Call logOffTEWindow("A","","PC5250","OK")
Call FinalStatus ()

'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [13,4062483]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


