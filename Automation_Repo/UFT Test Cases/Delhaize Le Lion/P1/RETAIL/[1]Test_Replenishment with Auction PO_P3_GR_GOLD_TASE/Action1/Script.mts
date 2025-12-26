
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Replenishment with Auction PO_P3_GR_GOLD_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Replenishment with Auction PO_P3_GR_GOLD_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\TASE_DT_06DCACCDB02_002_Create_Delivery_for_integrated_store_Fresh_Artemis.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2

Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''''--------------login----------------'''''

Call Launch_Mainframe(DT_ENVIRONMENT,"A","PC5250","OK")
Call Verify_Launch_LoginPage("A")
Call Login_Mainframe("A",DT_USERNAME,DT_PASSWORD)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY)
Call Verify_TEFieldExists("ENVGOLD","","A","","PC5250","&No","LogIN Page")
Call SendTEScreenKey ("A",DT_P_KEY)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C1)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID,DT_P_FIELD_VAL_SELECT_ENVIORINMENT,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C1,DT_P_FIELD_VAL_ENVIORINMENT_771_C1,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C1)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C2,DT_P_FIELD_VAL_ACTIVATE_ENVIORINMENT_C2,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C2)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C2)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C3,DT_P_FIELD_VAL_ACTIVATE_ENVIORINMENT_C3,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C3)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C4,DT_P_FIELD_VAL_MENU_INITIALISE_C4,"")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C4)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME,DT_P_FIELD_VAL_SELECT_MENU_GENERAL_PMS_C5)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C5)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField ("A",DT_P_FIELD_NAME_C1,DT_P_FIELD_VAL_SELECT_MENUS_DE_RECEPTIONS_C6)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C6)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField ("A",DT_P_FIELD_NAME_C2,DT_P_FIELD_VAL_RECEPTION_C7)
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C7)
'Call CaptureScreenshot_Mainframe("A")
'Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C5,DT_P_FIELD_VAL_PERS_CODE_C8,"")
'Call SendTEScreenKey ("A",DT_P_SEND_KEY_C7)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField ("A",DT_P_FIELD_NAME_C3,DT_P_FIELD_VAL_QUAI_NUMBER_C9)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C6,DT_P_FIELD_VAL_NUMERO_ORD_C10,"")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C7,DT_P_FIELD_VAL_BDL_FOURN_C11,DT_DESC)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C8)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C3)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C4)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C8,DT_P_FIELD_VAL_ARTICLE_C12,DT_DESC_C1)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C5)
Call CaptureScreenshot_Mainframe("A")
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C9,DT_P_FIELD_VAL_CHECK_PO_NUMBER_C13,DT_DESC_C2)
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C10,"DT_OP_DATE","")
Call CaptureScreenshot_Mainframe("A")
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTEField("A",DT_P_FIELD_NAME_C5,DT_P_FIELD_VAL_QUANTITY_C15)
Call SetTEField("A",DT_P_FIELD_NAME_C6,DT_P_FIELD_VAL_QUANTITY_C16)
Call CaptureScreenshot_Mainframe("A")
'''DT_P_FIELD_VAL_DLV_DATE_C171 = ConvertDoubledigit(CSTR(Day(DT_OP_DATE)))+""+ConvertDoubledigit( Cstr(Month(DT_OP_DATE))) + ""+ Cstr(Year(DT_OP_DATE))
DT_P_FIELD_VAL_DLV_DATE_C171 = ConvertDoubledigit(CSTR(Day(DT_OP_DATE)))+""+ConvertDoubledigit( Cstr(Month(DT_OP_DATE))) + ""+ Cstr(Year(DT_OP_DATE))
Call SetTEField("A",DT_P_FIELD_NAME_C7,(DT_P_FIELD_VAL_DLV_DATE_C171))
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C6)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C7)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C8)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C9)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField ("A",DT_P_FIELD_NAME_C4,DT_P_FIELD_VAL_C14)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C9)
Call CaptureScreenshot_Mainframe("A")
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


