
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from DC - CC RS01_p4
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_Returns from DC - CC RS01_p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Returns from DC - CC RS01_p4_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario


Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",DT_INCREMENT_REFERENCE+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''----------------------Tcode MIRO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,True)
Call PressEnter()
Call SetTextboxNolabel("INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
'Call SetTextbox("Posting Date",INVFO-BUDAT,"",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
'Call PressEnter()
Call SetComboByKey("RM08M-VORGANG", DT_MIRO_6000_TRANSACTION)
Call PressEnter()
Call TakeScreenShot

Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call PressEnter() 
Call TakeScreenShot
Call SetComboByKey("RM08M-ITEM_LIST_VERSION", DT_MIRO_6310_LAYOUT)
Call PressEnter() 
Call SetTextBoxNoLabel("RM08M-LFSNR",0,DT_MIRO_6212_RM08MLFSNR,False)
Call PressEnter()
Call TakeScreenShot

Call SelectRowGuiTable("SAPLMR1MTC_MSEL_VEN","PO",DT_PO,True)
Call ClickButton("Continue   \(F8\)",False)

Call SelectTab("HEADER","Details",False)
Call SetComboByKey("INVFO-BLART", DT_MIRO_0150_DOC_TYPE)
Call PressEnter()
Call TakeScreenShot

Call SelectTab("HEADER","Basic Data",False)
Call SelectCheckBox("INVFO-XMWST","0","ON",False)
Call TakeScreenShot
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT,False)
Call PressEnter() 
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",false)
Call GetStatusBar("item1","DT_MIRO_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIRO_STATUSBAR_OUTPUT&" created")

'-------------------------MIR4-----------------------
' SetTcode(tcode)
Call SetTcode(DT_MIRO_6000_OKCD)
Call Pressenter()
Call TakeScreenShot

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIRO_6150_INVOICE_DOCUMENT_NO,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIRO_6150_FISCAL_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",false)
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

