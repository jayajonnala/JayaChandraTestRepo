
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


gstrTestCaseName = "Test_Verify GI document_Frozen"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''' Login '''
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)
'
''INPUT''
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",1,DT_MIGO_2010_GODYNPRODOC_YEAR,FALSE)
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",1,DT_MIGO_2010_GODYNPROMAT_DOC,False) 
Call PressEnter() 
Call SelectTab("TS_GOITEM","Transfer Posting",False)
call VerifyTextBoxContent("Article","GODYNPRO-MAKTX",9,DT_MIGO_0390_CHECK_TEXT_OF_ARTICLE,false)
call VerifyTextBoxNoLabelContent("GODYNPRO-MATNR",8,DT_MIGO_0390_CHECK_TEXT_OF_GODYNPROMATNR,false)
Call SelectTab("TS_GOITEM","Where",False)
call VerifyTextBoxContent("Site","GOITEM-WERKS",9,DT_MIGO_0390_CHECK_TEXT_OF_SITE,false)   ''verifynolabelcontent
call VerifyTextBoxContent("Storage Location","GOITEM-LGORT",15,DT_MIGO_0390_CHECK_TEXT_OF_STOR_LOC,false)   ''verifynolabelcontent
call VerifyTextBoxContent("Goods recipient","GOITEM-WEMPF",13,DT_MIGO_0390_CHECK_TEXT_OF_GOITEMUMWRK,false)   ''verifynolabelcontent
call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
call VerifyTableCellContent(2,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
call VerifyTableCellContent(3,"Movement type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
call VerifyTableCellContent(1,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)
call VerifyTableCellContent(2,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_1)
call VerifyTableCellContent(3,"Direction","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_2)
call VerifyTableCellContent(1,"Purchase order","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_0)
call VerifyTableCellContent(2,"Purchase order","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_1)
call VerifyTableCellContent(3,"Purchase order","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_PURCHASE_ORDER_2)
call VerifyTableCellContent(1,"Item","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ITEM_0)
call VerifyTableCellContent(2,"Item","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ITEM_1)
call VerifyTableCellContent(3,"Item","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_ITEM_2)

Call LogOff() 
Call FinalStatus ()
