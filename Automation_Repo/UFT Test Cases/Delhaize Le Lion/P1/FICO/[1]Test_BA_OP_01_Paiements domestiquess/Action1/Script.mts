
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


gstrTestCaseName = "Test_BA_OP_01_Paiements domestiquess"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'' Login '''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)
Call TakeScreenShot

call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,true) 
Call TakeScreenShot
'call ClickButton("Continue   \(Enter\)",true) 
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
call ClickButtonIfExist("Continue   \(Enter\)",true)    
Call TakeScreenShot
call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,false) 
'This function SetTextboxNoLabel is used for vendor to supplier change.
call SetTextbox("Invoice date","INVFO-BLDAT","",convertdate(DT_FB60_0010_POSTING_DATE),false)  
call SetTextbox("Reference","INVFO-XBLNR","",DT_FB60_0010_REFERENCE,false)  
'call SetTextbox("Posting Date","INVFO-BUDAT","",convertdate(DT_FB60_0010_POSTING_DATE),false)  
Call TakeScreenShot
call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,false)  
call SetTextbox("Text","INVFO-SGTXT","",DT_FB60_0010_TEXT,false) 
Call SetTextbox("Tax Amount","INVFO-WMWST","",DT_FB60_0010_TAX_AMOUNT,False)
Call TakeScreenShot
call SetTableData("SAPLFSKBTABLE","G/L acct",1,"","",DT_FB60_0100_TABLECELL_GL_ACCT_0,false)
call SetTableData("SAPLFSKBTABLE","Amount in doc.curr.",1,"","",DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0,false)
call SetTableData("SAPLFSKBTABLE","Tax code",1,"","",DT_FB60_0100_TABLECELL_TAX_CODE_0,false)
call SetTableData("SAPLFSKBTABLE","Cost center",1,"","",DT_FB60_0100_TABLECELL_COST_CENTER_0,false)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyTextBoxContent("Bal\.","RF05A-AZSAL","",DT_FB60_1100_CHECK_TEXT_OF_BAL,False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1","DT_GENERATED_DOC_Output")
Call WriteRunTimeDataToExcelGlobalSheet("DT_GENERATED_DOC_Output",DT_GENERATED_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusBar(DT_GENERATED_DOC_1)
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)

'''Script to extract the crct payment date from 82 to 110 by KGARA on 14/12/22'''

'Call SetTcode(DT_FB60_0200_OKCD)     
'Call PressEnter()     
'Call TakeScreenShot
'Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB60_1000_VENDOR_ACCOUNT,false)
'Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FB60_1000_COMPANY_CODE_OCC1,false)
'Call SelectRadioButton("X_OPSEL","Open items",False)
'Call TakeScreenShot
'Call ClickButton("Execute   \(F8\)",False)
'Call TakeScreenShot
'Call 
'Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
'Call SelectRowGuiTable("SAPLSKBHTC_WRITE_LIST","Column content","Document number",True)
'Call ClickButtonIfExist("Hide Sel\. Fields \(Ctrl\+F2\)",True)
'Call ClickButtonIfExist("Copy   \(Enter\)",True)
''' SelectColumnGuiGrid(gridTitle, gridIndex, columnName, blnIsItPopup)
'Call SelectColumnGuiGrid("","","Document Number",False)
'Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)
'Call TakeScreenShot
'Call SetTextbox("Document Number","%%DYN001-LOW","",DT_GENERATED_DOC,True)
'Call TakeScreenShot
'Call ClickButton("Execute   \(Enter\)",true)
'Call TakeScreenShot
'Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
'' SelectMenuBar(pathName)
'Call SelectMenuBar("Settings;Switch list")
'Call ClickButtonToolBar("&FIND",0)
'Call SetTextbox("Search Term:","GS_SEARCH-VALUE","","Payment date",True)
'Call ClickButton("OK   \(Enter\)",True)
'Call TakeScreenShot
'Call ClickButton("Cancel   \(F12\)",True)
'Call ClickButton("Show Selected Fields \(F7\)",True)
'Call ClickButton("Transfer   \(Enter\)",True)
'Call TakeScreenShot
'Call GetGridContent("","","Payment date",1,"Document Number",DT_GENERATED_DOC,"DT_Pymt_dt_Output")
'Call WriteRunTimeDataToExcelGlobalSheet("DT_Pymt_dt_Output",DT_Pymt_dt)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''--------------f110---------------'''
Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Identification","F110V-LAUFI","",DT_FB60_0200_IDENTIFICATION,False)
Call SetTextbox("Run Date","F110V-LAUFD","",ConvertDate(DT_FB60_0200_RUN_DATE),False)
Call TakeScreenShot
Call SelectTab("F110_TABSTRIP","Parameter",False)
Call SetTextbox("Posting Date","F110C-BUDAT","",ConvertDate(DT_FB60_0202_POSTING_DATE),False)
Call TakeScreenShot
Call SetTextbox("Docs entered up to","F110C-GRDAT","",ConvertDate(DT_FB60_0202_DOCS_ENTERED_UP_TO),False)
Call SetTableData("SAPF110VCTRL_FKTTAB","Company codes","1","","",DT_FB60_0202_TABLECELL_COMPANY_CODES_0,False)
Call TakeScreenShot
Call SetTableData("SAPF110VCTRL_FKTTAB","Pmt meths","1","","",DT_FB60_0202_TABLECELL_PMT_METHS_0,False)
Call SetTableData("SAPF110VCTRL_FKTTAB","Next p/date","1","","",ConvertDate(DT_FB60_0202_TABLECELL_NEXT_PDATE_0),False)'''''DT_Pymt_dt,False)ConvertDate(DT_FB60_0202_TABLECELL_NEXT_PDATE_0),False)
Call TakeScreenShot
Call SetTextboxNoName("Supplier","",DT_FB60_7004_VENDOR,False)
Call TakeScreenShot
Call SelectTab("F110_TABSTRIP","Free selection",False)
Call TakeScreenShot
Call FocusTextBoxByIndex("Field Name","F110V-TEXT1","0",False)
'Call SendKey("{F4}")
'Call TakeScreenShot
'Call ClickButtonIfExist("All Values   \(Shift\+F6\)",True)
'Call SendKey("{F2}")
Call SetTextbox("Values","F110V-LIST1","",DT_FB60_0203_VALUES,False)
Call TakeScreenShot
Call SelectTab("F110_TABSTRIP","Additional Log",False)
Call SelectCheckBox("F110V-XTRFA","1","ON",False)
Call SelectCheckBox("F110V-XTRZW","1","ON",False)
Call SelectCheckBox("F110V-XTRBL","1","ON",False)
Call SetTextbox("Vendors \(from/to\)","F110V-VONKK","",DT_FB60_7004_VENDOR,false)
Call TakeScreenShot
Call SelectTab("F110_TABSTRIP","Status",False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot
Call ClickButton("Schedule Proposal   \(Shift\+F1\)",False)
Call SetTextbox("Start date","F110V-STRDT","",ConvertDate(DT_FB60_1106_START_DATE),True)
Call TakeScreenShot
Call SelectCheckbox("F110V-XSTRF","1","ON",True)
Call TakeScreenShot
Call ClickButton("Schedule   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Status   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Status   \(Shift\+F2\)",False)
Call ClickButton("Display Proposal   \(Shift\+F9\)",False)
Call TakeScreenShot
Call VerifyGridCellContentByRefColumn("",1,"Supplier","10001658","Ind: Incoming/Outgoing Payment","","Outgoing Payment")
Call VerifyGridCellContentByRefColumn("",1,"Supplier","10001658","Indicator: Payment/Exception","","Payment")
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call ClickButton("Schedule payment run   \(F7\)",False)
Call TakeScreenShot
Call SetTextbox("Start date","F110V-STRDT","",ConvertDate(DT_FB60_1106_START_DATE),True)
Call ClickButton("Schedule   \(Enter\)",True)
Call ClickButton("Status   \(Shift\+F2\)",False)
Call TakeScreenShot
''' VerifyTextBoxNoLabelContent(textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxNoLabelContent("F110V-STATU","3",DT_FB60_0201_CHECK_TEXT_OF_F110VSTATU,false)
''''''''''------------fbl1n-----------'''''''''''
Call SetTcode(DT_FB60_0200_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB60_1000_VENDOR_ACCOUNT,false)
Call SetTextbox("Company code","KD_BUKRS-LOW","",DT_FB60_1000_COMPANY_CODE_OCC1,false)
Call SelectRadioButton("X_AISEL","All items",False)
Call TakeScreenShot
Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call TakeScreenShot
Call SetTextbox("Document Number","%%DYN012-LOW","",DT_FB60_1106_DOCUMENT_NUMBER,False)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Change layout...   \(Ctrl\+F8\)",False)
Call SelectRowGuiTable("SAPLSKBHTC_WRITE_LIST","Column content","Document number",True)
Call ClickButtonIfExist("Hide Sel\. Fields \(Ctrl\+F2\)",True)
Call ClickButtonIfExist("Copy   \(Enter\)",True)
' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
Call VerifyifGuiLabelExistsByRelativeid(" S","wnd\[0\]/usr/lbl\[6,10\]")
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()


