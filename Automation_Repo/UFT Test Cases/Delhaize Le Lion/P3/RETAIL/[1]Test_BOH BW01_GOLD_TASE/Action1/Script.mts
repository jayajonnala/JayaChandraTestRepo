
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
'.................Test_BOH BW01_GOLD_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_BOH BW01_GOLD_TASE"
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

Call SendTEScreenKey ("A",DT_P_KEY)
Call SendTEScreenKey ("A",DT_P_KEY_C1)

Call Verify_TEFieldExists("ENVGOLD","","A","","PC5250","&No","LogIN Page")

Call SetTEFieldBY_ID ("A",DT_P_FIELDID,DT_P_FIELDVALUE,"")
Call SendTEScreenKey ("A",DT_P_SENDKEY)

Call CaptureScreenshot_Mainframe("A")

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C1,DT_P_FIELDVALUE_C1,"")
Call SendTEScreenKey ("A",DT_P_SENDKEY_C1)

Call CaptureScreenshot_Mainframe("A")

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C2,DT_P_FIELDVALUE_C2,"")
Call SendTEScreenKey ("A",DT_P_SENDKEY_C2)

Call SendTEScreenKey ("A",DT_P_KEY_C2)

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C3,DT_P_FIELDVALUE_C3,"")
Call SendTEScreenKey ("A",DT_P_SENDKEY_C3)
Call CaptureScreenshot_Mainframe("A")

Call SetTEFieldBY_ID ("A",DT_P_FIELDID_C4,DT_P_FIELDVALUE_C4,"")
Call SendTEScreenKey ("A",DT_P_SENDKEY_C4)

Call SetTEField("A",DT_P_FIELDNAME,DT_P_FIELDVALUE_C5)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C5)

Call CaptureScreenshot_Mainframe("A")

Call SetTEField("A",DT_P_FIELDNAME_C1,DT_P_FIELDVALUE_C6)
Call SendTEScreenKey ("A",DT_P_SENDKEY_C6)

Call CaptureScreenshot_Mainframe("A")

Call SetTEFieldBY_ID("A",DT_P_FIELDID_C5,DT_P_FIELDVALUE_C7,"")
Call SetTEFieldBY_ID("A",DT_P_FIELDID_C6,DT_P_FIELDVALUE_C8,"")

Call SendTEScreenKey ("A",DT_P_SENDKEY_C7)

Call CaptureScreenshot_Mainframe("A")

CAll VerifyTEFieldByID("A", DT_P_FIELDID_C7, DT_P_FIELDVALUE_C9, "")

Call SendTEScreenKey ("A",DT_P_KEY_C3)
Call CaptureScreenshot_Mainframe("A")

Call SendTEScreenKey ("A",DT_P_KEY_C4)
Call CaptureScreenshot_Mainframe("A")

Call SetTEFieldBY_ID("A",DT_P_FIELDID_C8,DT_P_FIELDVALUE_C10,DT_P_DESCRIPTION)
Call CaptureScreenshot_Mainframe("A")
Call SendTEScreenKey ("A",DT_P_SENDKEY_C8)

Call logOffTEWindow("A","","PC5250","OK")
Call FinalStatus ()

