

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
'.................Test Script Name : Test_02GR04_006_LocVendFranch_an_Deficit_Rev_Franch_auto_AB
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR04_006_LocVendFranch_an_Deficit_Rev_Franch_auto_AB"
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
Call SetComboByKey("GODYNPRO-REFDOC","R02")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call PressEnter() 
Wait(1)
Call TakeScreenShot()
Call ClickButton("Check Entries   \(F7\)",False)
Wait(1)
Call VerifyStatusBarMessageType("S")

Call VerifyStatusBar(DT_MIGO_1_CHECK_TEXT_OF_STATUSBAR)

Call ClickButtonIfExist("MIGO_OK_GO",False)
Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

'Post the Document
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")

'''------------------------------TCode WE9-------------------------------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIGO_1_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'
'Fill The Details and Execute
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 
Call SelectRowGuiGrid("Variant Catalog for Program RSEIDOC9"&".*","","Variant name",DT_MIGO_100_VARIANT,True)
Call ClickButtonIfExist("Choose   \(F2\)",True) 
Wait(2)
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_MIGO_1000_FOR_VALUE_,False)
Wait(3)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call ClickButtonIfExist("Yes",True)
Wait(5)
Call VerifyStatusBar("IDocs were found")

'Call VerifyifGuiLabelExists(DT_MIGO_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)
'Call GetLabelContentByRefLabel(Right(DT_MIGO_100_VARIANT_OCC1,Len(DT_MIGO_100_VARIANT_OCC1)),994,0,"DT_IDOC_NUMBER_OUTPUT",False)
Call GetLabelContentByRefLabel(Right(DT_MIGO_100_VARIANT_OCC1,Len(DT_MIGO_100_VARIANT_OCC1)),1136,0,"DT_IDOC_NUMBER_OUTPUT",False)
'Call GetLabelContentByRefLabel(Right(DT_MIGO_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK,Len(DT_MIGO_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)),497,0,"DT_IDOC_NUMBER_OUTPUT",False)
Call ClickLabel(DT_IDOC_NUMBER_OUTPUT,0,False)
Call ClickLinkGuiTree("IDoc "& DT_IDOC_NUMBER_OUTPUT&";Data records;#1;E1BP2017_GM_HEAD_01","E1BP2017_GM_HEAD_01","",False)
Call VerifyTableCellContent(7,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_MIGO_1000_FOR_VALUE_)
Call TakeScreenShot()


'##############################################################################################################
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

'################################################################################################################
