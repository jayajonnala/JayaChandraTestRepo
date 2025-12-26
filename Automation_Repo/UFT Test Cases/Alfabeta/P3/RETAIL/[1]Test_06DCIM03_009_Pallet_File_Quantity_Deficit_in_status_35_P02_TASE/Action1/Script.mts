
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
'................Test_06DCIM03_009_Pallet_File_Quantity_Deficit_in_status_35_P02
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_06DCIM03_009_Deficit_35_P02_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\Log_Artemis\TASE_DT_06-09-02-Full Gr on Sloc1 (goods +pallet).xls"
'gstrresultFolderPath="C:\Users\aprus\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''''--------------login----------------'''''


Call Launch_Mainframe(DT_ENVIRONMENT,"A","PC5250","OK")

If Dialog("text:=IBM i signon").WinEdit("attached text:=&Password:").Exist Then
	Dialog("text:=IBM i signon").WinEdit("attached text:=&Password:").Set DT_PASSWORD
	Dialog("text:=IBM i signon").WinButton("text:=OK").Click
	Wait 5
End If

Call Verify_Launch_LoginPage("A")

Call Login_Mainframe("A",DT_USERNAME,DT_PASSWORD)
Call SendTEScreenKey ("A",DT_P_SENDKEY)

CAll SetTEField("A", DT_P_FIELDNAME, DT_P_FIELDVALUE)
Call SendTEScreenKey ("A",DT_P_SENDKEY)

CAll SetTEField("A", DT_P_FIELDNAME_C1, DT_P_FIELDVALUE_C1)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C1)

CAll SetTEField("A", DT_P_FIELDNAME_C2, DT_P_FIELDVALUE_C2)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C2)

CAll SetTEField("A", DT_P_FIELDNAME_C3, DT_P_FIELDVALUE_C3)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID,DT_P_FIELDVALUE_C4,DT_P_DESCRIPTION)
Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C1,DT_P_FIELDVALUE_C5,DT_P_DESCRIPTION_C1)
Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C2,DT_P_FIELDVALUE_C6,DT_P_DESCRIPTION_C2)

Call SetTEFieldBY_ID ("A",DT_P_COLUMN,DT_P_VALUE,"")

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C3,DT_P_FIELDVALUE_C7,DT_P_DESCRIPTION_C3)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C3)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C4,DT_P_FIELDVALUE_C8,DT_P_DESCRIPTION_C4)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C4)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C5,DT_P_FIELDVALUE_C9,DT_P_DESCRIPTION_C5)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C5)

Call VerifyTEFieldByID("A", DT_P_FIELDID_C6, DT_P_FIELDVALUE_C10, DT_P_DESCRIPTION_C6)
Call VerifyTEFieldByID("A", DT_P_FIELDID_C7, DT_P_FIELDVALUE_C11, DT_P_DESCRIPTION_C7)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C8,DT_P_FIELDVALUE_C12,DT_P_DESCRIPTION_C8)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C6)

Call CaptureScreenshot_Mainframe("A")

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C9,DT_P_FIELDVALUE_C13,DT_P_DESCRIPTION_C9)

Call SendTEScreenKey ("A",DT_P_KEY_C1)

CAll Get_TEFieldValue_byID("A", DT_P_FIELDID_C9, "DT_PIECES_BEFORE_OUTPUT", DT_P_DESCRIPTION_C9)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C10,DT_P_FIELDVALUE_C13,DT_P_DESCRIPTION_C10)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C7)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C11,DT_P_FIELDVALUE_C14,DT_P_DESCRIPTION_C11)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C8)

CAll Get_TEFieldValue_byID("A", DT_P_FIELDID_C12, "DT_PIECES_OUTPUT", DT_P_DESCRIPTION_C12)

Call CaptureScreenshot_Mainframe("A")

Call SendTEScreenKey ("A",DT_P_KEY)

Call SendTEScreenKey ("A",DT_P_KEY_C1)

Call SendTEScreenKey ("A",DT_P_KEY_C2)

CAll Set_TEScreen("A", DT_P_ROW_C1, DT_P_COLUMN_C1, DT_P_VALUE_C1)
Call SendTEScreenKey ("A",DT_P_KEY_C3)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C13,DT_P_FIELDVALUE_C15,DT_P_DESCRIPTION_C13)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C14,DT_P_FIELDVALUE_C16,DT_P_DESCRIPTION_C14)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C9)

Call VerifyTEFieldByID("A", DT_P_FIELDID_C15, DT_P_FIELDVALUE_C17, DT_P_DESCRIPTION_C15)
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


