

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06PROD02_005_DC_Production_Fresh_Cancellation
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
	DataRowSet= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_06PROD02_005_Fresh_Cancellation"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


''--------------------------------MFBF-----------------------------

Call ClickButton("Doc\.-Specific Reversal of Posting   \(Shift\+F6\)",False)
Call TakeScreenShot()
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SelectRadioButton("RM61A-RTYPO","Article Document",False)
 Call SetTextbox("Article Document","RM07M-MBLNR","",DT_MFBF_0400_ARTICLE_DOCUMENT_OCC1,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
call ClickButtonIfExist("Continue   \(Enter\)",True)
Call PressEnter() 
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_DOC_ART_MOV_OUTPUT")
Call VerifyStatusBar("Article movement with document "& DT_DOC_ART_MOV_OUTPUT & " reversed")

Call ClickButton("Back   \(F3\)",False)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

''--------------------------------MF12-----------------------------
Call SetTcode(DT_MFBF_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MFBF_0100_OKCD)

call SetComboByKey("GODYNPRO-ACTION","A04")
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_DOC_ART_MOV_OUTPUT,False)
Call PressEnter()     ' 
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",False) 
Call GetGridContent("Documents in Accounting",0,"DOCNR",1,"<NA>","<NA>","DT_MFBF_0200_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")
Call SelectRowGuiGridbyRowNo("Documents in Accounting",0,1,True)
Call DoubleClickGuiGridCell("Documents in Accounting","",1,"Object type text",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'
Call LogOff()
Call FinalStatus ()


