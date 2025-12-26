'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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


''
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_MD_01_01_067-Update piece picking condition regarding to order UoM Orbit_p2_WE09_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 17th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_MD_01_01_067-Update piece picking condition regarding to order UoM Orbit_p2_WE09_TASE"


gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_MD_01_01_067-Update piece picking condition regarding to order UoM Orbit_p2_WE09_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 
'

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'''Increment the parameter/reload
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode WE09----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Created At","CRETIM-LOW","",DT_WE09_1000_CREATED_AT,False)
Call SetTextbox("to","CRETIM-HIGH","",DT_WE09_1000_TO,False)
Call SetTextbox("Created On","CREDAT-LOW","",DT_WE09_1000_CREATED_ON,False)
Call SetTextbox("to","CREDAT-HIGH","",DT_WE09_1000_TO_OCC1,False)
Call SetTextbox("Logical Message","MESTYP-LOW","",DT_WE09_1000_LOGICAL_MESSAGE,False)
Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_WE09_1000_SEARCH_IN_FIELD_,False)
Call SetTextbox("for Value \.\.\.","VALUE1_1","",DT_WE09_1000_FOR_VALUE_,False)
Call PressEnter()   
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Execute   \(F8\)",False)
wait(3)
Call ClickButtonIfExist("Yes",True)
wait(5)
'Capture the screenshot
Call TakeScreenShot()

Call ClickLabel("Time","",False)
Call ClickButtonIfExist("Sort in descending order   \(Ctrl\+Shift\+F4\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickLabel("Segment type","",False)
Wait(2)
Call SendKey("{DOWN}")
Wait(2)
Call SendKey("{F2}")
Wait(2)
'
I_DOC = SAPGuiSession("transaction:=WE09").SAPGuiWindow("transaction:=WE09").GetROProperty("text")
I_DOC = Replace(I_Doc,"Display: ","")
'
Call GetTextboxValue("EDIDC-STATUS","","DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Current Status","EDIDC-STATUS","",DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS,False)

'
''Updated Segment 000001 to E1KOMG for the property change
''Call ClickLinkGuiTree(I_DOC + ";Data records;E1KOMG","E1KOMG","",False)
Call ClickLinkGuiTree(I_DOC + ";Data records;Segment 000001","E1KOMG","",False)

Call GetTableCellData("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Cont.",8,"Fld Name","WERKS","DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_7_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(8,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_7)

'lickLink "IDoc 0000002190688999;Data records;E1KOMG;ZZMDPC_E1KOMG","ZZMDPC_E1KOMG"
'.ClickLink "IDoc 0000002190688999;Data records;E1KOMG;E1KONH","E1KONH"
'
Call ClickLinkGuiTree(I_DOC + ";Data records;Segment 000001;Segment 000002","ZZMDPC_E1KOMG","",False)
''Call ClickLinkGuiTree(I_DOC + ";Data records;E1KOMG;ZZMDPC_E1KOMG","ZZMDPC_E1KOMG","",False)
Call GetTableCellData("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Cont.",1,"Fld Name","BASEUOM","DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(1,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0)

Call ClickLinkGuiTree(I_DOC + ";Data records;Segment 000001;Segment 000003","E1KONH","",False)
''Call ClickLinkGuiTree(I_DOC + ";Data records;E1KOMG;E1KONH","E1KONH","",False)
Call GetTableCellData("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Cont.",2,"Fld Name","DATAB","DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(2,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_1)

Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)
Call SendKey("{DOWN}")
Wait(2)
Call SendKey("{F2}")
Wait(2)

I_DOC = SAPGuiSession("transaction:=WE09").SAPGuiWindow("transaction:=WE09").GetROProperty("text")
I_DOC = Replace(I_Doc,"Display: ","")

Call GetTextboxValue("EDIDC-STATUS","","DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS_OCC1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTextBoxContent("Current Status","EDIDC-STATUS","",DT_WE09_0100_CHECK_TEXT_OF_CURRENT_STATUS_OCC1,False)

Call ClickLinkGuiTree(I_DOC + ";Data records;Segment 000001","E1KOMG","",False)
''Call ClickLinkGuiTree(I_DOC + ";Data records;E1KOMG","E1KOMG","",False)
Call GetTableCellData("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Cont.",8,"Fld Name","WERKS","DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_7_OCC1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(8,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_7_OCC1)

Call ClickLinkGuiTree(I_DOC + ";Data records;Segment 000001;Segment 000002","ZZMDPC_E1KOMG","",False)
''Call ClickLinkGuiTree(I_DOC + ";Data records;E1KOMG;ZZMDPC_E1KOMG","ZZMDPC_E1KOMG","",False)
Call GetTableCellData("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Cont.",1,"Fld Name","BASEUOM","DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0_OCC1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(1,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_0_OCC1)

Call ClickLinkGuiTree(I_DOC + ";Data records;Segment 000001;Segment 000003","E1KONH","",False)
''Call ClickLinkGuiTree(I_DOC + ";Data records;E1KOMG;E1KONH","E1KONH","",False)
Call GetTableCellData("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Cont.",2,"Fld Name","DATAB","DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_1_OCC1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyTableCellContent(2,"Fld Cont.","IDOC_TREE_CONTROLINT_SEG_CONTROL",DT_WE09_0100_CHECK_TEXT_OF_TABLECELL_FLD_CONT_1_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()



