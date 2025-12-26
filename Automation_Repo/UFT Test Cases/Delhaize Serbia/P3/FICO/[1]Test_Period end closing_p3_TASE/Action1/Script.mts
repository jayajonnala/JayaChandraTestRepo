
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Period end closing_p3
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Period end closing_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Period end closing_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode KB34N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("ReverseDoc","COHEADER-BELNR_REV","",DT_KB34N_2100_REVERSEDOC,FALSE)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

'Click on Post
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_REV_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Document is posted under number "&DT_REV_DOCUMENT_NO_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectMenuBar("Posting;Display")

Call GetTextboxValue("COHEADER-BELNR", "", "DT_KB34N_2100_CHECK_TEXT_OF_DOCUMENTNO_OUTPUT", False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("DocumentNo","COHEADER-BELNR", "", DT_KB34N_2100_CHECK_TEXT_OF_DOCUMENTNO_OUTPUT, False)

'Verify table data
Call VerifyTableCellContent(1,"Rec. CCtr","SAPLSTC1GENERIC_TABLE_1",DT_KB34N_0100_CHECK_TEXT_OF_TABLECELL_REC_CCTR_0)
Call VerifyTableCellContent(1,"StatKF","SAPLSTC1GENERIC_TABLE_1",DT_KB34N_0100_CHECK_TEXT_OF_TABLECELL_STATKF_0)
Call VerifyTableCellContent(1,"Total Quantity","SAPLSTC1GENERIC_TABLE_1",DT_KB34N_0100_CHECK_TEXT_OF_TABLECELL_TOTAL_QUANTITY_0)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

