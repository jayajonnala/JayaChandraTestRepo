

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_V_AB_007_ Manage Vendor Lifecycle-Create Vendor (Direct Foreign EU no Subrange_Retail)
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

gstrTestCaseName = "Test_MD_V_AB_007_WITH_TAX_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  XK01------------------------------------------------
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT1",(DT_INCREMENT1)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT2",(DT_INCREMENT2)+1)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT3",(DT_INCREMENT3)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Account Group","RF02K-KTOKK","",DT_XK01_0100_ACCOUNT_GROUP,False)
Call SetTextbox("Company Code","RF02K-BUKRS","",DT_XK01_0100_COMPANY_CODE,False)
Call SetTextbox("PurchasingOrganization","RF02K-EKORG","",DT_XK01_0100_PURCHASINGORGANIZATION,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetCombo("SZA1_D0100-TITLE_MEDI",DT_XK01_0301_TITLE)
Call SetTextbox("Name","ADDR1_DATA-NAME1","",DT_XK01_0301_NAME,False)
Call SetTextbox("Search term 1/2","ADDR1_DATA-SORT1","",DT_XK01_0301_SEARCH_TERM_12,False)
Call SetTextbox("Street/House number","ADDR1_DATA-STREET","",DT_XK01_0301_STREETHOUSE_NUMBER,False)
Call SetTextbox("Street/House number","ADDR1_DATA-HOUSE_NUM1","",DT_XK01_0301_STREETHOUSE_NUMBER_OCC1,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-POST_CODE1","",DT_XK01_0301_POSTAL_CODECITY,False)
Call SetTextbox("Postal Code/City","ADDR1_DATA-CITY1","",DT_XK01_0301_POSTAL_CODECITY_OCC1,False)
Call SetTextbox("Country","ADDR1_DATA-COUNTRY","",DT_XK01_0301_COUNTRY,False)
Call SetTextbox("Telephone","SZA1_D0100-TEL_NUMBER","",DT_XK01_0301_TELEPHONE,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Tax Number 1","LFA1-STCD1","",FormatBlank(DT_XK01_0120_TAX_NUMBER_1),False)
Call TakeScreenShot()
Call SetTextbox("VAT Reg. No.","LFA1-STCEG","",FormatBlank(DT_XK01_0120_VAT_REG_NO),False) ' --- Getting error message as VAT Registration is already used by other Vendor .. Hence commented

Call FocusTextBox("Type of Busines","LFA1-J_1KFTBUS", False)
Call SendKey("{F4}")
wait 5
Call ClickButton("Find   \(Ctrl\+F\)",True)
Call SetTextbox("Find","RSYSF-STRING","",DT_XK01_0800_FIND,True)
Call SelectCheckbox("SCAN_STRING-START",0,DT_XK01_0800_STARTING_AT_CURRENT_LINE,True)
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)
Call SetFocusGuiLabel(DT_XK01_0800_FIND,"","",True)
Call  ClickButtonIfExist("Continue   \(Enter\)",True)
Call  ClickButtonIfExist("Copy   \(Enter\)",True)

Call SetTextbox("Tax office","LFA1-FISKU","",DT_XK01_0120_TAX_OFFICE,False)
Call SetTextbox("Location no\. 1","LFA1-BBBNR","",DT_XK01_0120_LOCATION_NO_1,False)
Call SetTextbox("Location no\. 2","LFA1-BBSNR","",DT_XK01_0120_LOCATION_NO_2,False)
Call SetTextbox("Check digit","LFA1-BUBKZ","",DT_XK01_0120_CHECK_DIGIT,False)
Call TakeScreenShot()
Call SetTextbox("Type of Industr","LFA1-J_1KFTIND","",FormatBlank(DT_TYPE_OF_INDUSTRY),False)
Call TakeScreenShot()

Call PressEnter()
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 
Call PressEnter() 


Call SetTextbox("Recon\. account","LFB1-AKONT","",DT_XK01_0210_RECON_ACCOUNT,False)

Call SetTextbox("Sort key","LFB1-ZUAWA","",DT_XK01_0210_SORT_KEY,False)
Call SetTextbox("Prev\.acct no\.","LFB1-ALTKN","",DT_XK01_0210_PREVACCT_NO,False)
Call SetTextbox("Cash mgmnt group","LFB1-FDGRV","",DT_XK01_0210_CASH_MGMNT_GROUP,False)
''This line is added based on the workshop corresnponding to defect 18916
Call SetTextbox("Release group","LFB1-FRGRP","",DT_XK01_0210_RELEASE_GROUP,False)
Call TakeScreenShot()
Call PressEnter()

Call SelectCheckbox("LFB1-REPRF",0,DT_XK01_0215_CHK_DOUBLE_INV,False)
Call SetTextboxNoLabel("LFB1-ZTERM","",DT_XK01_0215_PAYT_TERMS,False)
Call SetTextbox("Cr memo terms","LFB1-GUZTE","",DT_XK01_0215_CR_MEMO_TERMS,False)
Call SetTextbox("Payment methods","LFB1-ZWELS","",DT_XK01_0215_PAYMENT_METHODS,False)
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()
Call PressEnter()

Call SelectCheckbox("LFM1-WEBRE",0,DT_XK01_0310_GRBASED_INV_VERIF,False)
''This line is commented based on the workshop corresnponding to defect 18916
''Call SelectCheckbox("LFM1-XERSY",0,DT_XK01_0310_AUTOEVALGRSETMT_DEL,False)
Call SelectCheckbox("LFM1-BOLRE",0,DT_XK01_0310_SUBSEQUENT_SETTLEMENT,False)
Call SelectCheckbox("LFM1-BOIND",0,DT_XK01_0310_SUBSEQ_SETT_INDEX,False)
Call SelectCheckbox("LFM1-XNBWY",0,DT_XK01_0310_REVALUATION_ALLOWED,False)
Call SelectCheckbox("LFM1-NRGEW",0,DT_XK01_0310_GRANT_DISCOUNT_IN_KIND,False)
Call SelectCheckbox("LFM1-PRFRE",0,DT_XK01_0310_RELEVANT_FOR_PRICE_DET__VENDOR_HIERARCHY,False)
Call SelectCheckbox("LFM1-AGREL",0,DT_XK01_0310_RELEVANT_FOR_AGENCY_BUSINESS,False)

Call SetTextbox("Order currency","LFM1-WAERS","",DT_XK01_0310_ORDER_CURRENCY,False)
Call SetTextbox("Terms of paymnt","LFM1-ZTERM","",DT_XK01_0310_TERMS_OF_PAYMNT,False)
Call SetTextbox("Purchasing group","LFM1-EKGRP","",DT_XK01_0310_PURCHASING_GROUP,False)
Call SetTextbox("Planned deliv\. time","LFM1-PLIFZ","",DT_XK01_0310_PLANNED_DELIV_TIME,False)
''This line is added based on the workshop corresnponding to defect 18916
Call SetTextbox("Pricing Date Control","LFM1-MEPRF","",DT_XK01_0310_PRICING_DATE_CNTRL,False)
Call TakeScreenShot()
Call PressEnter()

Call SelectMenuBar("Environment;Classification")
Call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class",1,"","",DT_XK01_1600_TABLECELL_CLASS_0,False)
Call PressEnter()
Call SetTableData("SAPLCTMSCHARS_S","Value",1,"","",DT_XK01_5100_TABLECELL_VALUE_0,False)
Call SetTableData("SAPLCTMSCHARS_S","Value",2,"","",DT_XK01_5100_TABLECELL_VALUE_1,False)
Call SetTableData("SAPLCTMSCHARS_S","Value",8,"","",DT_MODE_OF_TRANSPORT,False)
Call PressEnter()
Call TakeScreenShot()
'Call SetVerticalScrollBar(10,False)
wait 5
Call SetTableData("SAPLCTMSCHARS_S","Value",22,"","",DT_XK01_5100_TABLECELL_VALUE_1_OCC1,False)
Call PressEnter()
Call TakeScreenShot()

Call FocusTableCell("SAPLCTMSCHARS_S","Value",23,"","","",False)
Call SetTableData("SAPLCTMSCHARS_S","Value",23,"","",DT_XK01_5100_TABLECELL_VALUE_2,False)
''This line is added based on the workshop corresnponding to defect 18916
''Call SetTableData("SAPLCTMSCHARS_S","Value",26,"","",DT_XK01_5100_TABLECELL_VALUE_5,False)
Call TakeScreenShot()
Call FocusTableCell("SAPLCTMSCHARS_S","Value",27,"","","",False)
Call SetTableData("SAPLCTMSCHARS_S","Value",27,"","",DT_XK01_5100_TABLECELL_VALUE_6,False)

Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Legacy Vendor ID .*",False)
Call SetTableData("SAPLZMDVM_LEGACY_VENDOR_ADDONSGTC_0100","Legacy Vendor Nr",1,"","",DT_XK01_0100_TABLECELL_LEGACY_VENDOR_NR_0,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Vendor "&DT_XK01_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" has been created for company code "& DT_XK01_0100_COMPANY_CODE &" purchasing organization "& DT_XK01_0100_PURCHASINGORGANIZATION)
'
Call LogOff()
Call FinalStatus ()

'Public Function FocusTableCell(tableName,columnName,rowNumber,refColumnName,refCellValue,cellValue,blnIsItPopup)
'	If Not (Environment.Value("blnFatalError") or cellValue= DS_SKIP) Then
'		If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : FocusTableCell"
'
'
'		Dim objTable,objWindow
'
'		strStepName ="Focus " & "'" & cellValue & "'"& " in Table '"&tableName&"'"
'    	Set objWindow= SetSAPwindowObj(blnIsItPopup)
'		Set objTable= SAPGuisession(sessionObject).sapguiwindow(objWindow).sapguitable(guiTable,"name:="&tableName)
'		If  objTable.Exist Then
'				
'				On Error Resume Next
'					If  rowNumber= "" Then
'						strRowNumber = objTable.FindRowByCellContent(refColumnName,refCellValue)
'					Else 
'						strRowNumber=rowNumber
'					End If
'
'					IF Err.Number = 13 Then
'							Call ReporterFunction(strLibraryFileName,"FocusTableCell","1","Table Data Field","refCellValue Not Found Within RefColumn. Check the --refCellValue--Parameter of--FocusTableCellByRefColumn-- Function Call")
'							strMsg = "'" & refCellValue & "'" & " Not Found Within reference column: '" & refColumnName & "' .Check the --refCellValue--Parameter of --FocusTableCell-- Function Call"
'							strStatus = "FAIL"
'							
'					ElseIf Err.Number=-2147220983 Then
'							Call ReporterFunction(strLibraryFileName,"FocusTableCell","1","Table Data Field","refColumn Not Found. Check the --refColumnName--Parameter of--FocusTableCellByRefColumn-- Function Call")
'							strMsg = "'" & refColumn & "'" & " Not Found. Check the --refColumnName--Parameter of --FocusTableCell-- Function Call"
'							strStatus = "FAIL"
'							blnObjectError=True
'
'                    End If 
'					If Err.Number = 0 Then
'						Err.Clear	
'						If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'					       ImagePath=CaptureScreenshot(strStepName,objTable,FALSE,False,TRUE)
'					    End if	
'
'					    strStepName ="Focus " & "Row Number: "& "'" & rowNumber & "'" & " and in Column Name: "& "'" & columnName& "'" & " in Table"
'					    On Error Resume Next
'					    strCellVal=objTable.GetCellData(rowNumber,columnName)
'
'					    objTable.SetCellData strRowNumber,columnName,strCellVal
'
'					If Err.Number = 0 Then
'						    Call ReporterFunction(strLibraryFileName,"FocusTableCell","2","Table Data Field","Value Used :-" & vbNewLine &"Table Name:" & tableName &vbNewLine&"Column Field: "& columnName & vbNewLine &"Row Field: "&strRowNumber & vbNewLine &" ValueField: "&strcellValue)
'						    strStatus = "DONE"
'						    strMsg = "Table Data set successfully"
'						    
'					ElseIf Err.Number=-2147220983  Then
'                            Call ReporterFunction(strLibraryFileName,"FocusTableCell","1","Table Data Field","Column Not Found. Check the --columnName--Parameter of--FocusTableCell-- Function Call")
'							strMsg = "Column Not Found. Check the --columnName--Parameter of --FocusTableCell-- Function Call"
'							strStatus = "FAIL"
'							blnObjectError=True
'					ElseIf Err.Number=-2147220975  Then
'						If not rowNumber="" Then
'						Call ReporterFunction(strLibraryFileName,"FocusTableCell","1","Table Data Field"," Row Not Found. Check the --rowNumber--Parameter of --FocusTableCell-- Function Call")
'							strMsg = "Row Not Found. Check the --rowNumber--Parameter of --FocusTableCell-- Function Call"
'							strStatus = "FAIL"
'						End If
'
'						End If	
'				 End If 
'					
'
'					Else 
'							Call ReporterFunction(strLibraryFileName,"FocusTableCell","1","Table Data Field","Table Object Missing")
'							strMsg = "Table Not Found- Please Check the Screen"
'							strStatus = "FAIL"
'							blnObjectError=True
'				End If
'
'
'	If  blnObjectError  Then
'		Environment.Value("blnFatalError")=True
'	End If
'
'	If strStatus = "FAIL"  Then
'		FocusTableCell = strMsg
'		blnMainFailFlag = True
'		ImagePath=CaptureScreenshot(strStepName,objTable,FALSE,False,TRUE)
'	Else
'		FocusTableCell = True
'	End If
'	If blnDefault_eSwiftReporting Then  
'		Call UpdateResultHtml(strStepName,cellValue,strMsg,strStatus,"")
'	End If
'End If
'
'Set strCellVal=Nothing
'Set objTable=Nothing
'Set objWindow=Nothing
'End Function   


