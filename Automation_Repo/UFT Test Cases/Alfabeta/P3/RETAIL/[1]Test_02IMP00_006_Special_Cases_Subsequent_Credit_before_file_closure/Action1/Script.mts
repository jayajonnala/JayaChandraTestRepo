	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_02IMP00_006_Special_Cases_Subsequent_Credit_before_file_closure"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_02IMP00_007_P5.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_DOCUMENT_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False) 
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call TakeScreenShot
call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
call SetComboByKey("RM08M-VORGANG",DT_MIRO_6000_TRANSACTION_OCC3)
Call TakeScreenShot
Call SetTextboxNoLabel("RM08M-EBELN","",DT_MIRO_6211_RM08MEBELN,False)
Call PressEnter()
Call TakeScreenShot
Call Clickbutton("Select All",False)
Call TakeScreenShot

Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",1,"ON",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",2,"ON",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Final Invoice",3,"ON",False)
Call TakeScreenShot

Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",1,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",2,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax Code",3,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call TakeScreenShot

Call SelectTab("HEADER","Details",False)
Call TakeScreenShot
Call SetTextbox("Unpl\. Del\. Csts","INVFO-BEZNK","",DT_MIRO_0010_AMOUNT,False)
Call SetComboByKey("Doc. Type",DT_MIRO_0150_DOC_TYPE_OCC2)
Call SelectTab("HEADER","Basic Data",False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot
Call GetStatusBar("item1","DT_Document_no_Output")

Call ClickButton("Other Invoice Document   \(Ctrl\+F10\)",False)
Call TakeScreenShot
call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)


'''''''''''''''we09''''''''''''''''''
Call SetTcode(DT_MIRO_6000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot
Call SetTextbox("Variant","V-LOW","",DT_MIRO_100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("for Value \.\.\.","VALUE1_1","",DT_MIRO_1000_FOR_VALUE_,False)
Call ClickButton("Execute   \(F8\)",false)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot
Call SetFocusGuiLabel("4",-21,0,false)
Call SendKey("{F2}")
Call TakeScreenShot
Call SetFocusGuiLabel("000000.*",32,72,False)
Call SendKey("{F2}")

Call ActivateNodeGuiTree("","#1;#2;#2;#1")
Call TakeScreenShot
''''''''''''''''mir4'''''''''''''''''
Call SetTcode(DT_MIRO_100_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6000_TEXT_OF_STATUSBAR,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False)
Call TakeScreenShot
Call GetGridContentByRefColumn("Documents in Accounting","0","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_Output")
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)
'''''''''''''''''''''nj1gp01'''''''''''''''''
Call SetTcode(DT_MIRO_6000_OKCD_OCC2) 
Call PressEnter()     ' 
Call TakeScreenShot


Call SetTextbox("Company Code","S_BUKRS","","GR02",False)
Call SetTextbox("Fiscal Year","S_GJAHR","",DT_YEAR,False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_1000_DOCUMENT_OCC2,False)
Call SelectCheckbox("S_TEST","0","OFF",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_MIRO_120_CHECK_TEXT_OF_STATUSBAR_Output")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
call VerifyStatusBar(DT_MIRO_120_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot


Call LogOff()
Call FinalStatus ()
