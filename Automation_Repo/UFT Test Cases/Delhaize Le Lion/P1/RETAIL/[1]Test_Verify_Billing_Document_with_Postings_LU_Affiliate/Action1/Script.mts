
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_Ver_Bill_Doc_Post_LU_Affi"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

''INPUT''
call SetTextbox("Billing document","VBRK-VBELN","",DT_VF03_0101_BILLING_DOCUMENT,false) 
Call PressEnter() 

call VerifyTableCellContent(1,"Article","SAPMV60ATCTRL_UEB_FAKT",DT_VF03_0104_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
call VerifyTableCellContent(2,"Article","SAPMV60ATCTRL_UEB_FAKT",DT_VF03_0104_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
call VerifyTableCellContent(1,"Billed Quantity","SAPMV60ATCTRL_UEB_FAKT",DT_VF03_0104_CHECK_TEXT_OF_TABLECELL_BILLED_QUANTITY_0)
'call VerifyTableCellContent(2,"Billed Quantity","SAPMV60ATCTRL_UEB_FAKT",DT_VF03_0104_CHECK_TEXT_OF_TABLECELL_BILLED_QUANTITY_1)
call ClickButton("Accounting overview   \(Shift\+F4\)",false)  
CALL ActivateCellGuiGridByRefVal("Documents in Accounting",0,"object type text","Accounting Document","Doc. Number",TRUE)
TakeScreenShot()

call VerifyGridCellContent("",1,"Account",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'call VerifyGridCellContent("",4,"Account",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
call VerifyGridCellContent("",5,"Account",0,DT_VF03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)

call ClickButton("Exit   \(Shift\+F3\)",false)  

Call LogOff()
Call FinalStatus ()
