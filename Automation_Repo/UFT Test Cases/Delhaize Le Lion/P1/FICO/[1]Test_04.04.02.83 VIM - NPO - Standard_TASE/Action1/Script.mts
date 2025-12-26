
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.83 VIM - NPO - Standard
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


gstrTestCaseName = "Test_04.04.02.83 VIM - NPO - Standard"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'all LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

'''--------------------------------------------  /n/OPT/VIM_7AX2----------------------------------------------

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call SetTextbox("Addl docs","S_DOCID","",DT_OPTVIM_7AX2_0002_ADDL_DOCS,False)
Call TakeScreenShot()
Call ClickButton("Copy From",False)
Call TakeScreenShot()

Call SetTextbox("Vendor Reference","/OPT/VIM_BL_1RIDX_OCR_DATA-XBLNR","",FormatBlank(DT_OPTVIM_7AX2_2002_VENDOR_REFERENCE),False)
Call SetTextbox("Gross Inv Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-GROSS_AMOUNT","",FormatBlank(DT_OPTVIM_7AX2_2002_GROSS_INV_AMOUNT),False)
Call SetTextbox("Expense Type","/OPT/VIM_BL_1RIDX_OCR_DATA-EXPENSE_TYPE","",FormatBlank(DT_OPTVIM_7AX2_2002_EXPENSE_TYPE),False)
Call SetTextbox("Requestor E-mail","/OPT/VIM_BL_1RIDX_OCR_DATA-EMAIL_ID","",FormatBlank(DT_OPTVIM_7AX2_2002_REQUESTOR_EMAIL),False)
Call SetTextbox("Requester ID","/OPT/VIM_BL_1RIDX_OCR_DATA-REQUISITIONER","",FormatBlank(DT_OPTVIM_7AX2_2002_REQUESTER_ID),False)
Call SetTextbox("VAT Amount","/OPT/VIM_BL_1RIDX_OCR_DATA-VAT_AMOUNT","",FormatBlank(DT_OPTVIM_7AX2_2002_VAT_AMOUNT),False)
Call SetTextbox("Tax Amount 1","/OPT/VIM_BL_1RIDX_OCR_DATA-TAXAMT_1","",FormatBlank(DT_OPTVIM_7AX2_2002_TAX_AMOUNT_1),False)
Call TakeScreenShot()

Call SetTextbox("Tax Rate 1","/OPT/VIM_BL_1RIDX_OCR_DATA-TAXRATE_1","",FormatBlank(DT_TAX_RATE1),False)
Call SetTextbox("PO Number","/OPT/VIM_BL_1RIDX_OCR_DATA-EBELN","",FormatBlank(DT_PO_NUMBER),False)
Call SetTextbox("Plant","/OPT/VIM_BL_1RIDX_OCR_DATA-WERKS","",FormatBlank(DT_PLANT),False)
Call SetTextboxNoLabel("/OPT/VIM_BL_1RIDX_OCR_DATA-TAX_RATE",0,FormatBlank(DT_TAX_RATE),False)
Call SetTextboxNoLabel("/OPT/VIM_BL_1RIDX_OCR_DATA-TAX_CODE",0,FormatBlank(DT_TAX_CODE),False)
Call TakeScreenShot()

Call SelectRowGuiTableByRow("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS",1,False)
Call ClickButton("Delete line",False)
Call TakeScreenShot()

Call  SetTableData("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS","G/L account",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_GL_ACCOUNT_0,False)
Call  SetTableData("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS","Cost Center",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_COST_CENTER_0,False)
Call  SetTableData("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS","Amount",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_AMOUNT_01,False)
Call SetTextbox("Arch\.","S_ARCHIV_ID","",DT_OPTVIM_7AX2_0002_ARCH,False)
Call  SetTableData("/OPT/VIM_BL_M1_SUBSCREENSTC_INDEX_ITEMS","Credit / Debit",1,"","",DT_OPTVIM_7AX2_0019_TABLECELL_CREDIT__DEBIT_0,False)

Call TakeScreenShot()
Call ClickButton("SUBMIT   \(F8\)",False)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call GetStatusBar("item2","DT_DOC_NO_OUTPUT")
Call GetStatusBar("item4","DT_WF_ID_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_OPTVIM_7AX2_0002_CHECK_TEXT_OF_STATUSBAR_OCC2)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()
