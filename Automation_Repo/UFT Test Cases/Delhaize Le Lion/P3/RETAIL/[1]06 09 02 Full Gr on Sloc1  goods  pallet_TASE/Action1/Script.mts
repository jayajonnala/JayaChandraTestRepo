
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
'.................06 09 02 Full Gr on Sloc1  goods  pallet_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "06 09 02 Full Gr Sloc1goodspallet"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\Log_Artemis\TASE_DT_06-09-02-Full Gr on Sloc1 (goods +pallet).xls"
'gstrresultFolderPath="C:\Users\aprus\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution1(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''''--------------login----------------'''''

Call Launch_Mainframe(DT_ENVIRONMENT,"A","PC5250","OK")
Call Verify_Launch_LoginPage("A")
Call Login_Mainframe("A",DT_USERNAME,DT_PASSWORD)
Call SendTEScreenKey ("A",DT_P_KEY)
Call CaptureScreenshot_Mainframe("A")
Call Verify_TEFieldExists("ENVGOLD","","A","","PC5250","&No","LogIN Page")
Call SendTEScreenKey ("A",DT_P_KEY)
Call SendTEScreenKey ("A",DT_P_KEY_C1)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID,DT_P_FIELD_ID_VAL_CHOICE,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C1,DT_P_FIELD_ID_VAL_CHOICE_C1,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C1)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C2,DT_P_FIELD_ID_VAL_CHOICE_C2,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C2)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C2)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C3,DT_P_FIELD_ID_VAL_CHOICE_C3,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C3)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C4,DT_P_FIELD_ID_VAL_CHOICE_C4,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C4)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME,DT_P_FIELD_ID_VAL_CHOICE_C5)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C5)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C1,DT_P_FIELD_ID_VAL_CHOICE_C6)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C6)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A",DT_P_FIELD_NAME_C2,DT_P_FIELD_ID_VAL_CHOICE_C7)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C7)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C5,DT_P_FIELD_ID_VAL_PALLET_C8,"")
Call SetTEField("A",DT_P_FIELD_NAME_C3,DT_P_FIELD_ID_VAL_QUAI_C9)
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C6,DT_P_FIELD_ID_VAL_PO_C10,"")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C7,DT_P_FIELD_ID_VAL_PO_C11,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C8)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C3)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C4)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C8,DT_P_FIELD_ID_VAL_ARTICLE_C12,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C9)
Call VerifyTEFieldByID ("A",DT_P_FIELD_ID_C9,DT_P_FIELD_ID_VAL_QUANTITY_ARTICLE_C13,DT_DESC)
Call SetTEField("A","field1044",DT_P_FIELD_ID_QUANTITY_C21)
Call SetTEField("A","field1049",DT_P_FIELD_ID_QUANTITY_C22)
Call CaptureScreenshot_Mainframe("A")
Call Get_TEFieldValue_byID("A",DT_P_FIELD_ID_C10,"DT_OP_GET_DATE","")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTEField("A",DT_P_FIELD_NAME_C4,ConvertDateTxt(DT_OP_GET_DATE,DT_DATE_INCREMENT))
Call SetTEField("A",DT_P_FIELD_NAME_C9,DT_P_FIELD_ID_VAL_QUANTITY_C26)
Call SetTEField("A",DT_P_FIELD_NAME_C10,DT_P_FIELD_ID_VAL_QUANTITY_C27)
Call SetTEField("A",DT_P_FIELD_NAME_C11,ConvertDateTxt(DT_OP_GET_DATE,DT_DATE_INCREMENT))
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C10)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C8)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C8,DT_P_FIELD_ID_VAL_ARTICLE_C12,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C9)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C10)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C10)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C11,DT_P_FIELD_ID_VAL_ARTICLE_C15,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C9)
Call CaptureScreenshot_Mainframe("A")
Call SetTEField("A","field1044",2)
Call SetTEField("A","field1049",80)
Call SetTEField("A",DT_P_FIELD_NAME_C4,ConvertDateTxt(DT_OP_GET_DATE,DT_DATE_INCREMENT))
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C10)
Call CaptureScreenshot_Mainframe("A")
Call SetTEFieldBY_ID ("A",DT_P_FIELD_ID_C11,DT_P_FIELD_ID_VAL_ARTICLE_C15,"")
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SEND_KEY_C9)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C11)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C11)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C12)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_KEY_C12)
Call CaptureScreenshot_Mainframe("A")
Call logOffTEWindow("A","","PC5250","OK")
Call CaptureScreenshot_Mainframe("A")
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


