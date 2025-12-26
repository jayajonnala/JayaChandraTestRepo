
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


gstrTestCaseName = "Test_Verify_Billing_Document"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'---------------------vl03n--------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,false)
Call PressEnter() 
call ClickButton("Document Flow   \(F7\)",false)  
Call TakeScreenShot
If ValidateNodeTextGuiTree(0,"Invoice") = True Then
	Call ActivateNodeGuiTree(0,"#1;#1;RegExp:=Inv.*")
	Call GetGridContent(".*", 0, "Doc.no.", 1, "Item","", "DT_VL03N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
	Call TakeScreenShot()

Else
	Call ActivateNodeGuiTree(0,"#1;#1;RegExp:=Inter.*")
	Call GetGridContent(".*", 0, "Doc.no.", 1, "Item","", "DT_VL03N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
       Call TakeScreenShot()	
End If
''' ActivateItemGuiTree(treeIndex, itemPath, itemText)
'Call SelectNodeGuiTree(0,"#1;#1;RegExp:=Intercompany .*")
''Call SelectNodeGuiTree(0,"#1;#1;RegExp:=Invoice .*")
'Call DoubleClick()
'Call TakeScreenShot
'call GetGridContent(".*",0,"Doc.no.",1,"NA","NA","DT_VL03N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")
call ClickButton("Back   \(F3\)",false)
call ClickButton("Back   \(F3\)",false)
call ClickButton("Back   \(F3\)",false)
call ClickButton("Back   \(F3\)",false)


'---------------------------------VF03------------------

Call SetTcode(DT_VL03N_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_VL03N_0100_OKCD)

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Billing document","VBRK-VBELN","",DT_VL03N_0101_BILLING_DOCUMENT,False) 
Call pressenter()
call VerifyTableCellContent(1,"Description","SAPMV60ATCTRL_UEB_FAKT",UCase(DT_VL03N_0104_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_0))
call VerifyTableCellContent(2,"Description","SAPMV60ATCTRL_UEB_FAKT",Ucase(DT_VL03N_0104_CHECK_TEXT_OF_TABLECELL_DESCRIPTION_1))
''MSGBOX DT_VL03N_0104_BILLED_QUANTITY_0
Call TakeScreenShot
call VerifyTableCellContent(1,"Billed Quantity","SAPMV60ATCTRL_UEB_FAKT",DT_QUANTITY_0)
call VerifyTableCellContent(2,"Billed Quantity","SAPMV60ATCTRL_UEB_FAKT",DT_QUANTITY_1)
'
Call TakeScreenShot
Call LogOff()
Call FinalStatus ()


