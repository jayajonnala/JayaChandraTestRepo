
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_Return to Vendor BW08_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Return to Vendor BW08_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\Log_Artemis\TASE_DT_06-09-02-Full Gr on Sloc1 (goods +pallet).xls"
'gstrresultFolderPath="C:\Users\aprus\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''''--------------login----------------'''''

Call Launch_Mainframe(DT_ENVIRONMENT,"A","PC5250","OK")
Call Verify_Launch_LoginPage("A")
Call Login_Mainframe("A",DT_USERNAME,DT_PASSWORD)
Call SendTEScreenKey ("A",DT_P_KEY)
Call Verify_TEFieldExists("ENVGOLD","","A","","PC5250","&No","LogIN Page")
Call SendTEScreenKey ("A",DT_P_KEY)
Call SendTEScreenKey ("A",DT_P_KEY_C1)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID,DT_P_FIELD_ID_VAL_ENV_CHOICE,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C2)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C1,DT_P_FIELD_ID_VAL_BW08_C1,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C3)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C2,DT_P_FIELD_ID_VAL_MGENERAL_C2,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY)
Call SendTEScreenKey ("A",DT_P_KEY_C4)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C3,DT_P_FIELD_ID_VAL_MGENERAL_C3,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C1)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C4,DT_P_FIELD_ID_VAL_CHOIX_C4,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C2)
Call SetTEField("A",DT_P_FIELD_NAME,DT_P_FIELD_ID_VAL_OPTION_C5)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C5)
Call SetTEField("A",DT_P_FIELD_NAME_C1,DT_P_FIELD_ID_VAL_OPTION_C6)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C6)
Call SetTEField("A",DT_P_FIELD_NAME_C2,DT_P_FIELD_ID_VAL_OPTION_C7)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C7)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C5,DT_P_FIELD_ID_VAL_CODE_PERS_C8,"")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C6,DT_P_FIELD_ID_VAL_CODE_FOURN_C9,"")
Call SendTEScreenKey ("A",DT_P_KEY_C8)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C7,DT_P_FIELD_ID_VAL_PALLET_NUMBER_C10,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C3)
Call CaptureScreenshot_Mainframe("A")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C8,"DT_OP_GET_CODE","")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C9,"DT_OP_GET_CODE_C1","")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C10,"DT_OP_GET_UV_C2","")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C11,"DT_OP_GET_UV_C3","")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C12,"DT_OP_GET_PALLET_C4","")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C13,"DT_OP_GET_PALLET_C5","")
DataRowSet=2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C14,DT_P_FIELD_ID_VAL_VERIFY_C11,"")
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C15,DT_P_FIELD_ID_VAL_VERIFY_C12,"")
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C16,DT_P_FIELD_ID_VAL_VERIFY_C13,"")
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C17,DT_P_FIELD_ID_VAL_VERIFY_C14,"")
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C18,DT_P_FIELD_ID_VAL_VERIFY_C15,"")
Call SendTEScreenKey ("A",DT_P_KEY_C9)
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


