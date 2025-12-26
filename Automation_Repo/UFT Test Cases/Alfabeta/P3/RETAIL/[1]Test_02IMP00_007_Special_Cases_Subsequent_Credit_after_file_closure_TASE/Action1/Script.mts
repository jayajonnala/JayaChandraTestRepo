

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

gstrTestCaseName = "Test_02IMP00_007_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_02IMP00_007_P5.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

'//-----------------------------------MIGO -----------------------------------

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_MIRO_1000_COMPANY_CODE)
Call TakeScreenShot()
call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot
call SetComboByKey("Transaction",3)
Call TakeScreenShot
Call SetTextboxNoLabel("INVFO-BLDAT","",Replace(DT_MIRO_0010_DOCUMENT_DATE,"/","."),False)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)     ' - Line (19)
Call PressEnter()
Call TakeScreenShot
call SetComboByKey("RM08M-REFERENZBELEGTYP",1)
Call TakeScreenShot
call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot

call SetComboByKey("Transaction",DT_MIRO_6000_TRANSACTION)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call TakeScreenShot
Call SetComboByKey("RM08M-XWARE_BNK",DT_MIRO_6211_RM08MXWARE_BNK)
Call SetTextboxNoLabel("RM08M-EBELN","",DT_MIRO_6211_RM08MEBELN,False)
Call PressEnter()
Call TakeScreenShot
Call SelectRowGuiTable("SAPLMR1MTC_MSEL_VEN","Vendor","ALPHA BANK",True)
Call ClickButtonIfExist("Continue   \(F8\)",True)
Call SetCombo("Tax Amount","A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)")
Call TakeScreenShot

' SetTableDataNoRef(tableName, columnName, rowNumber, cellValue, blnIsItPopup)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax code",1,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax code",2,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax code",3,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax code",4,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","Tax code",5,"A0 (GR VAT 0,0% Domestic Purchases - Trade Goods)",False)


Call SelectTab("HEADER","Details",False)
Call TakeScreenShot
Call SetTextbox("Unpl\. Del\. Csts","INVFO-BEZNK","",DT_MIRO_0150_UNPL_DEL_CSTS,False)
Call SetComboByKey("Doc. Type",DT_MIRO_0150_DOC_TYPE)

Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Select All",False)
Call SelectMenuBar("Invoice Document;Post")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot
Call GetStatusBar("item1","DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_OUTPUT")



'''''''''''''''we09''''''''''''''''''
Call SetTcode(DT_MIRO_6000_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Variant","V-LOW","",DT_MIRO_100_VARIANT,True)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextBox("for Value \.\.\.","VALUE1_1","",DT_MIRO_1000_FOR_VALUE_,False)
Call ClickButton("Execute   \(F8\)",false)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call TakeScreenShot
Call SetFocusGuiLabel("4",-21,0,false)
Call SendKey("{F2}")
Call TakeScreenShot
Call SetFocusGuiLabel("00000018.*",21,0,False)
Call SendKey("{F2}")
Call ActivateNodeGuiTree("","#1;#2;#2;#1")
Call TakeScreenShot
'''''''''''''''mir4'''''''''''''''''
Call SetTcode(DT_MIRO_100_OKCD) 
Call PressEnter()     ' 
Call TakeScreenShot

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6000_TEXT_OF_STATUSBAR,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(DATE),False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetGridContentByRefColumn("Documents in Accounting","0","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_OF_GRIDCELL_0_DOCNR_Output")
Call TakeScreenShot
Call ClickButton("Cancel   \(F12\)",True)

'''''''''''''''''''''nj1gp01'''''''''''''''''

Call SetTcode(DT_MIRO_6000_OKCD_OCC2) 
Call PressEnter()     ' 
Call TakeScreenShot

'''Call SetTextbox("Company Code","S_BUKRS","","GR02",False)
'''Call SetTextbox("Fiscal Year","S_GJAHR","","2022",False)
Call SetTextbox("Company Code","S_BUKRS","",DT_J1GP01_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","S_GJAHR","",DT_J1GP01_FISCAL_YEAR,False)
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_1000_DOCUMENT,False)
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

