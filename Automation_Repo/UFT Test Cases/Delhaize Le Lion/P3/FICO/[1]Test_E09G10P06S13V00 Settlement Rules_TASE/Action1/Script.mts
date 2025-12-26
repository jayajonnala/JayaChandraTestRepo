'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_E09G10P06S13V00 Settlement Rules  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_E09G10P06S13V00 Settlement Rules"
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
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
'
''Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-ZFIGL_UPLOAD_POST----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("File path name","P_FILE","",DT_ZFIGL_UPLOAD_POST_1000_FILE_PATH_NAME,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
wait 10
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButton("Yes",True)
Call TakeScreenShot

''''''--------TransactionCode--SM35---------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Sess.:","D0100-MAPN","",DT_ZFIGL_UPLOAD_POST_1005_SESS,False)
Call SetTextbox("From:","D0100-VON","",ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM),False)
Call SetTextbox("Created by:","D0100-CREATOR","",DT_ZFIGL_UPLOAD_POST_1005_CREATED_BY,False)
Call TakeScreenShot
Call PressEnter()

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Process session   \(F8\)",False)
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call TakeScreenShot

Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
Call TakeScreenShot

Call SetTextbox("From:","D0100-VON","",ConvertDate(DT_ZFIGL_UPLOAD_POST_1005_FROM_OCC1),False)
Call TakeScreenShot
Call PressEnter()
Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Log   \(F7\)",False)
Call TakeScreenShot
Call SelectRowGuiTableByRow("RSBDC_PROTOCOLTC_PROTOCOL", 1, False)
Call ClickButton("Display   \(F2\)",False)
Call TakeScreenShot

Call GetTableCellData("RSBDC_PROTOCOLTC_PROT_DIS", "Message", 2, "", "", "DT_OP_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",4)

Call VerifyTableCellContent(2, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_2_OCC1)


Call VerifyTableCellContent(4, "Message", "RSBDC_PROTOCOLTC_PROT_DIS",DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4)
Call VerifyTableCellContent(5, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_5)
Call VerifyTableCellContent(6, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_6)
Call VerifyTableCellContent(7, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7)

'''Call VerifyTableCellContent(4, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", Cstr(DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4))
'''Call VerifyTableCellContent(5, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", Cstr(DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_5))
'''Call VerifyTableCellContent(6, "Message", "RSBDC_PROTOCOLTC_PROT_DIS",Cstr( DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_6))
'''Call VerifyTableCellContent(7, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", Cstr(DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7))
'''
'''
''Call VerifyTableCellContent(4, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_4)
''Call VerifyTableCellContent(5, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_5)
''Call VerifyTableCellContent(6, "Message", "RSBDC_PROTOCOLTC_PROT_DIS", DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_6)
''Call VerifyTableCellContent(7, "Message", "RSBDC_PROTOCOLTC_PROT_DIS",DT_ZFIGL_UPLOAD_POST_1400_CHECK_TEXT_OF_TABLECELL_MESSAGE_7)
''
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot

'''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SetTextbox("Account Number","RACCT-LOW","",DT_ZFIGL_UPLOAD_POST_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_ZFIGL_UPLOAD_POST_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot


Call DoubleClickGuiGridCell("","",Month(date())+1,"Debit",False)
Call TakeScreenShot

Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False) 
Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ZFIGL_UPLOAD_POST_1105_DOCUMENT_NUMBER,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

'Call VerifyifGuiLabelExists(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
''' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Document Number","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Amount in local currency","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)

' VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)

''Call VerifyGridCellContentbyName("",1,"Document Number","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
''Call VerifyGridCellContentbyName("",1,"Amount in local currency","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
'Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB, 1)

Call ClickButton("Change Layout...   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","*wbs*",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonifExist("Last Column   \(Ctrl\+F4\)",False)
'Call VerifyifGuiLabelExists_ByIndex("A01332.02.BZ000131-P2.0E",0)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"WBS Element","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROJK)


''CAll VerifyGridCellContentbyName("",1,"WBS Element","","A01332.02.BZ000131-P2.0E")
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot

'''''''--------TransactionCode-/CJ8G----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC8)

Call SetTextbox("Controlling Area","SVALD-VALUE","",DT_ZFIGL_UPLOAD_POST_0300_CONTROLLING_AREA,False)
Call PressEnter()     
Call TakeScreenShot
Call SelectCheckbox("LKO74-TESTLAUF",0, "OFF", False)
Call SetTextbox("Fiscal Year","LKO74-GJAHR","",DT_ZFIGL_UPLOAD_POST_1000_FISCAL_YEAR,False)
Call SetTextbox("Selection variant","PRZB-VARIANT","",DT_ZFIGL_UPLOAD_POST_0500_SELECTION_VARIANT,False)
Call SetTextbox("Settlement period","LKO74-PERIO","",DT_ZFIGL_UPLOAD_POST_1000_SETTLEMENT_PERIOD,False)  
Call SetTextbox("Posting period","LKO74-BUPERIO","",DT_ZFIGL_UPLOAD_POST_1000_POSTING_PERIOD,False)           
'Call SetCombo("LKO74-VAART",DT_ZFIGL_UPLOAD_POST_1000_PROCESSING_TYPE)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Back   \(F3\)",False)
Call ClickButtonIfExist("Back   \(F3\)",False)
Call TakeScreenShot

''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_ZFIGL_UPLOAD_POST_0100_OKCD_OCC3)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC10)

Call SetTextbox("Account Number","RACCT-LOW","",DT_ZFIGL_UPLOAD_POST_1000_TO,False)
Call SetTextbox("to","RACCT-HIGH","",DT_ZFIGL_UPLOAD_POST_1000_ACCOUNT_NUMBER_OCC1,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_ZFIGL_UPLOAD_POST_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_ZFIGL_UPLOAD_POST_1000_FISCAL_YEAR_OCC1,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

''Need to do it Generic
Call DoubleClickGuiGridCell("","",MOnth(Date())+1,"Period",False)
Call TakeScreenShot
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call ClickButtonToolBar("&FIND",0)

Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","DOCUMENT NUMBER",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call ClickButton("Define Filter Values",True)
Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ZFIGL_UPLOAD_POST_1105_DOCUMENT_NUMBER_OCC1,True)
Call SetTextbox("to","%%DYN001-HIGH","",DT_ZFIGL_UPLOAD_POST_1105_TO,True)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot

'Call SetTextbox("Find","GD_SEARCHSTR","","DOCUMENT NUMBER",True)
'Call TakeScreenShot
'Call ClickButton("Continue   \(Enter\)",True)
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_ZFIGL_UPLOAD_POST_1105_DOCUMENT_NUMBER_OCC1,True)
'Call SetTextbox("to","%%DYN001-HIGH","",DT_ZFIGL_UPLOAD_POST_1105_TO,True)
'Call TakeScreenShot
'Call ClickButton("Execute   \(Enter\)",True)
'Call TakeScreenShot

Call ClickButton("Change Layout...   \(Ctrl\+F8\)",False)
Call TakeScreenShot
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","*wbs*",True)
'Call SetTextbox("Find","GD_SEARCHSTR","","*wbs*",True)
Call TakeScreenShot
Call ClickButton("OK   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
Call TakeScreenShot
Call ClickButton("Show Selected Fields \(F7\)",True)
Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot


'
'Call ClickButton("Continue   \(Enter\)",True)
'Call ClickButton("Show sel. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)
'Call TakeScreenShot

' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContent("",1,"Document Number","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
'Call VerifyGridCellContent("",2,"Document Number","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BELNR)

Call VerifyGridCellContent("",1,"WBS element","", DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PROJK_OCC1)
'Call VerifyGridCellContent("",2,"WBS element","", DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PROJK)

Call VerifyGridCellContent("",1,"Amount in local currency","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
'Call VerifyGridCellContent("",2,"Amount in local currency","",DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_DMSHB)



'''
''Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_1105_TO,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_1105_DOCUMENT_NUMBER_OCC1,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB,0)
''
''Call ClickButton("Last Column   \(Ctrl\+F4\)",False)
''Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PROJK,0)
''Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PROJK,1)
''
''Call ClickButton("Last Page   \(Ctrl\+Page down\)",False)
''Call VerifyifGuiLabelExists_ByIndex(DT_ZFIGL_UPLOAD_POST_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_DMSHB,0)
''
''
'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

