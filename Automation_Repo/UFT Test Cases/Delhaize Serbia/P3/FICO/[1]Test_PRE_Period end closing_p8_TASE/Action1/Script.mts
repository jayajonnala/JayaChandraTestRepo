
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_Period end closing_p8
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
'.................Test Script Name :Test_PRE_Period end closing_p8
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_PRE_Period end closing_p8_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode KB31N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter Details
Call SetTextbox("Doc\. Date","COHEADER-BLDAT","",ConvertDate(DT_KB31N_2100_DOC_DATE),FALSE)
Call SetTextbox("Postg Date","COHEADER-BUDAT","",ConvertDate(DT_KB31N_2100_POSTG_DATE),FALSE)
Call SetTextbox("Doc\. Text","COHEADER-BLTXT","",DT_KB31N_2100_DOC_TEXT,FALSE)
Call TakeScreenShot()


Call SetTableData("SAPLSTC1GENERIC_TABLE_1","Rec. CCtr","1","","",DT_KB31N_0100_TABLECELL_REC_CCTR_0,False)
Call SetTableData("SAPLSTC1GENERIC_TABLE_1","StatKF","1","","",DT_KB31N_0100_TABLECELL_STATKF_0,False)
Call SetTableData("SAPLSTC1GENERIC_TABLE_1","Total Quantity","1","","",DT_KB31N_0100_TABLECELL_TOTAL_QUANTITY_0,False)
Call TakeScreenShot()
Call PressEnter()

'Get the period Value
Call GetTextboxValue("COHEADER-PERIO",0,"DT_Period",False)

'Click on Post
Call ClickButton("Post   \(Ctrl\+S\)",False)

Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Document is posted under number "&DT_DOCUMENT_NO_OUTPUT)

Call SelectMenuBar("Posting;Display")

''Get the Document No
Call GetTextboxValue("COHEADER-BELNR",0,"DT_Document_No",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyTextBoxContent("DocumentNo","COHEADER-BELNR",0,DT_KB31N_2100_CHECK_TEXT_OF_DOCUMENTNO_OCC1,False)
Call VerifyTableCellContent(1,"Rec. CCtr","SAPLSTC1GENERIC_TABLE_1",DT_KB31N_0100_CHECK_TEXT_OF_TABLECELL_REC_CCTR_0)
Call VerifyTableCellContent(1,"StatKF","SAPLSTC1GENERIC_TABLE_1",DT_KB31N_0100_CHECK_TEXT_OF_TABLECELL_STATKF_0)
Call VerifyTableCellContent(1,"Total Quantity","SAPLSTC1GENERIC_TABLE_1",DT_KB31N_0100_CHECK_TEXT_OF_TABLECELL_TOTAL_QUANTITY_0)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

