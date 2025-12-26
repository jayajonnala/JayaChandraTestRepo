

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
'.................Test Script Name : Test_02GR09_001_Reversal_Loc_Vend_Franch_w_Trading_Goods
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR09_001_Reversal_Loc_Vend_Franch_w_Trading_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR08_001_Reversal_Loc_Vend_DSD_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'''Login'''
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  


'----------------------Tcode MIGO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the PO Number and Press Enter
Call SetCombo("GODYNPRO-ACTION","Cancellation")
 ' SetComboByKey(attachedTextOrComboName, keyValue)
Call SetComboByKey("GODYNPRO-REFDOC","R02")
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call PressEnter() 
Wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("MIGO_OK_GO",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

Call ClickButton("Next Item",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

Call ClickButton("Next Item",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()


'Validate the Movement Type
Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE)

Call ClickButton("Check Entries   \(F7\)",False)
Wait(1)


Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")

''------------------------------Display Article Document Details-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetCombo("GODYNPRO-ACTION","Display")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ARTICLE_DOCUMENT,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,2021,False)
Call PressEnter() 
Wait(2)
Call TakeScreenShot()


'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Check Screen
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'Click Display Document Header
Call ClickButton("Display Document Header   \(F5\)",False)
Wait(1)
Call TakeScreenShot()


'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
wait(2)

''------------------------------TCode WE9-------------------------------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIGO_750_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'
'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9"&".*","","Variant name","/MBGMCR_OUT",True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Wait(2)
Call SetTextboxNoLabel("VALUE1_1",0,DT_MIGO_1000_FOR_VALUE__OCC2,False)
Wait(3)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call ClickButtonIfExist("Yes",True)
Wait(5)
Call VerifyStatusBar("IDocs were found")

'Call GetLabelContentByRefLabel(Right(DT_MIGO_100_VARIANT,Len(DT_MIGO_100_VARIANT)),994,0,"DT_IDOC_NUMBER",False)
Call GetLabelContentByRefLabel(Right(DT_MIGO_100_VARIANT,Len(DT_MIGO_100_VARIANT)),1136,0,"DT_IDOC_NUMBER",False)
'Call ClickLabel(DT_IDOC_NUMBER,0,False)
'Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
'Call VerifyTableCellContent(7,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_MIGO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_6)
'

If VerifyifGuiLabelCondition(DT_ME21N_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK) = True Then
	
	Call VerifyifGuiLabelExists(DT_MIGO_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)
	Call ClickLabel(DT_IDOC_NUMBER,0,False)
	Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
	Call VerifyTableCellContent(7,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_MIGO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_6)
	Call TakeScreenShot()

	
Else

	Call SetTcode(DT_BD87_OKCD)
	Call TakeScreenShot()
	Call PressEnter() 
	
	Call SetTextbox("IDoc Number","SX_DOCNU-LOW","",DT_IDOC_NUMBER,False)
	Call TakeScreenShot()
	Call ClickButton("Execute   \(F8\)",False)
	Call TakeScreenShot()
	Call SelectItemGuiTree(0,"Retail Pre-Production;IDocs in outbound processing;IDoc ready for dispatch (ALE service)","IDoc ready for dispatch (ALE service)")
	Call TakeScreenShot()
	Call ClickButton("Process Selected Node   \(F8\)",False)
	Call TakeScreenShot()
	Call ClickButton("Continue   \(Enter\)",True)
	Call TakeScreenShot()
	Call VerifyGridCellContent("Processed IDocs", 1, "Status text", 0, DT_ME21N_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)


	Call SetTcode(DT_MIGO_750_OKCD) 
	Call PressEnter() 
	Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'Fill The Details and Execute
	Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
	Call SetTextbox("Created By","ENAME-LOW","","",True)
	Call ClickButtonIfExist("Execute   \(F8\)",True) 
	Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9.*","","Variant name","/MBGMCR_OUT",True)
	Call ClickButtonIfExist("Choose   \(F2\)",True) 
	Call SetTextboxNoLabel("VALUE1_1",0,DT_MIGO_1000_FOR_VALUE__OCC2,False)
	Wait(3)
	Call TakeScreenShot()
	Call ClickButton("Execute   \(F8\)",False) 
	Wait(2)
	Call ClickButtonIfExist("Yes",True)
	Wait(5)
	Call VerifyStatusBar("IDocs were found")

	'Call GetLabelContentByRefLabel(Right(DT_MIGO_100_VARIANT,Len(DT_MIGO_100_VARIANT)),994,0,"DT_IDOC_NUMBER",False)
	Call GetLabelContentByRefLabel(Right(DT_MIGO_100_VARIANT,Len(DT_MIGO_100_VARIANT)),1136,0,"DT_IDOC_NUMBER",False)
	Call VerifyifGuiLabelExists(DT_MIGO_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)
	Call ClickLabel(DT_IDOC_NUMBER,0,False)
	Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
	Call VerifyTableCellContent(7,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_MIGO_100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_6)
	Call TakeScreenShot()

End If



'##############################################################################################################
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

'################################################################################################################
