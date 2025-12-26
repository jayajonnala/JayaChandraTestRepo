'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G09P01S12V32 REVERSE AA Document 
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G09P01S12V32 REVERSE AA"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''''--------TransactionCode-AB08----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","RLAB01-BUKRS","",DT_AB08_0010_COMPANY_CODE,False)
Call SetTextbox("Asset","RLAB01-ANLN1","",DT_AB08_0010_ASSET,False)
Call SetTextbox("Sub-number","RLAB01-ANLN2","",DT_AB08_0010_SUBNUMBER,False) 
Call SetTextbox("Fiscal Year","RLAB01-GJAHR","",DT_AB08_0010_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()

Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100",2,False)
Call ClickButton("Reverse Document   \(F6\)",False)
Call TakeScreenShot

Call SetTextbox("Posting Date","BAPI6037_REV_DATA-PSTNG_DATE","",ConvertDate(DT_AB08_0650_POSTING_DATE),True)
Call SetTextbox("Period","BAPI6037_REV_DATA-FIS_PERIOD","",DT_AB08_0650_PERIOD,True)
Call SetTextbox("Reversal Reason","BAPI6037_REV_DATA-REASON_REV","",DT_AB08_0650_REVERSAL_REASON,True)
Call ClickButton("Continue   \(Enter\)",True)
Call SelectRowGuiGridbyRowNo("Line items", "", 2, False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_OP_DOC")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call VerifyStatusBar(DT_AB08_0010_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)

'''''--------TransactionCode-AW01----------''''

Call SetTcode(DT_AB08_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Change","",DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",1,"Year-end","",DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_JENDE)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Change","",DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation",4,"Year-end","",DT_AB08_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)
Call TakeScreenShot
Call DoubleClickGuiGridCell("Transactions", "",2, "Amount posted", False)

''''''--------TransactionCode-FB03----------''''

Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot

Call GetGridContentByTitle("", "", "Account", 2, DT_OP_AB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call TakeScreenShot

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",datatable_row)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_AB08_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)


Call SetTextbox("Account Number","RACCT-LOW","",DT_AB08_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_AB08_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_AB08_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call DoubleClickGuiGridCell("",1,MOnth(Date)+1,"Period",False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","","Document Number",True)
Call ClickButton("Continue   \(Enter\)",True)
Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_AB08_1105_DOCUMENT_NUMBER,True)
Call SetTextbox("to","%%DYN001-HIGH","",DT_AB08_1105_TO,True)
Call ClickButton("Execute   \(Enter\)",True) 

'Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'Call ClickButtonToolBar("&FIND", True)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Document Number",True)
'Call SelectCheckbox("GS_SEARCH-EXACT_WORD",0, "ON", True)
'Call ClickButton("OK   \(Enter\)",True)
'Call ClickButton("Cancel   \(F12\)",True)
'Call ClickButton("Add Filter Criterion \(F7\)",True)
'Call TakeScreenShot          
'Call ClickButton("Define Filter Values",True)
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_AB08_1105_DOCUMENT_NUMBER,True)
'Call SetTextbox("to","%%DYN001-HIGH","",DT_AB08_1105_TO,True)
'Call ClickButton("Execute   \(Enter\)",True)  

Call TakeScreenShot

' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
Call VerifyifGuiLabelExistsByRelativeid(DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SEGMENT,"wnd\[0\]/usr/lbl\[95,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SEGMENT,"wnd\[0\]/usr/lbl\[95,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[84,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR,"wnd\[0\]/usr/lbl\[84,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[43,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB,"wnd\[0\]/usr/lbl\[43,9\]")

'Call VerifyGridCellContent("", 1, "Segment", "", DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SEGMENT)
'Call VerifyGridCellContent("", 2, "Segment", "", DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SEGMENT)
'Call VerifyGridCellContent("", 1, "Profit Center", "", DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'Call VerifyGridCellContent("", 2, "Profit Center", "", DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
'Call VerifyGridCellContent("", 1, "Amount in local currency", "", DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
'Call VerifyGridCellContent("", 2, "Amount in local currency", "", DT_AB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call TakeScreenShot
Call LogOff()
Call FinalStatus ()
'*********************************************End Of Script*********************************************************************
