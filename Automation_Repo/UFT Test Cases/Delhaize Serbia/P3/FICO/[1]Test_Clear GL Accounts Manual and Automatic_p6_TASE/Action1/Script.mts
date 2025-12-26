'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Clear GL Accounts Manual and Automatic_p6_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th April
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Clear GL Accounts Manual and Automatic_p6_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p6_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''----------------------Tcode FB50----------------------------
'Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()
	
'this step is not in the log but it is a mandate in screen
'Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB50_1000_COMPANY_CODE,True)
Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_FB50_1000_COMPANY_CODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Continue   \(Enter\)",True)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Reference","ACGL_HEAD-XBLNR","",DT_FB50_1010_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","ACGL_HEAD-BKTXT","",DT_FB50_1010_DOCHEADER_TEXT,False)
Call SetTextbox("Document Date","ACGL_HEAD-BLDAT","",Replace((DT_FB50_1010_DOCUMENT_DATE),"/","."),False)
Call SetTextbox("Posting Date","ACGL_HEAD-BUDAT","",Replace((DT_FB50_1010_POSTING_DATE),"/","."),False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document type","ACGL_HEAD-BLART","",DT_SETDOCTYPE,False)

Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",1,DT_FB50_0100_TABLECELL_GL_ACCT_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","G/L acct",2,DT_FB50_0100_TABLECELL_GL_ACCT_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","D/C",1,DT_FB50_0100_TABLECELL_DC_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","D/C",2,DT_FB50_0100_TABLECELL_DC_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",1,DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Amount in doc.curr.",2,DT_FB50_0100_TABLECELL_AMOUNT_IN_DOCCURR_1,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Profit center",1,DT_FB50_0100_TABLECELL_PC_0,False)
Call SetTableDataNoRef("SAPLFSKBTABLE","Profit center",2,DT_FB50_0100_TABLECELL_PC_1,False)

Call PressEnter() 
wait(1)
'Capture the screenshot
Call TakeScreenShot()

'Click on Post Buton
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If document is posted and get the status bar nummber
Call GetStatusBar("item1","DT_FB50_1001_DOCUMENT_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusBarMessageType("S")
VerifyStatusBar(DT_FB50_1001_CHECK_TEXT_OF_STATUSBAR)
'Capture the screenshot
Call TakeScreenShot()


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

