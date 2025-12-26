		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.15.01 Itinerary_Movement Type F218
'.................Author : TCS 
'................ Creation Date :
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If


gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.06.05.15.01 Itinerary_Movement Type F218_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''--------TransactionCode-FAGLL03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)
DT_FAGLL03_1000_POSTING_DATE= "01.01."& CSTR(Year(Date))
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FAGLL03_1000_TO),False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot

Call SetHorizontalScrollBar(75,False)
Call TakeScreenshot

Call ClickLabel("Text", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Text","%%DYN001-LOW","",DT_FAGLL03_1105_TEXT,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call SetHorizontalScrollBar(0,False)
Call VerifyifGuiLabelExists(DT_FAGLL03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_)
Call TakeScreenshot
Call LogOff'
Call FinalStatus()
