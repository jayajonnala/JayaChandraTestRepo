
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Import - regular goods_p2
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Import - regular goods_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Import - regular goods_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Import - regular goods_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME23N----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Click on Other purchase order
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call TakeScreenShot()

Call ClickButtonIfExist("Other Document   \(Enter\)",True)
wait(2)
Call TakeScreenShot()

'Click on Messages button and get the output type.
Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
wait(2)

'Verify the Status
Call VerifyTableCellContent(3,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_2)
Call VerifyTableCellContent(4,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_3)

'Verify the Output Type data
Call VerifyTableCellContent(3,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)
Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)

''----------------------Tcode ME9F----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME23N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_0100_OKCD)

'Enter the details
Call SetTextbox("Document Number","S_EBELN-LOW","",DT_ME23N_1000_DOCUMENT_NUMBER,False)   
Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME23N_1000_PURCHASING_ORGANIZATION,False)   
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_ME23N_1000_PURCHASING_GROUP,False)   
Call SetTextbox("Application","P_KAPPL","",DT_ME23N_1000_APPLICATION,False)   
Call SetTextbox("Processing Status","P_VSTAT","",DT_ME23N_1000_PROCESSING_STATUS,False)  
Call TakeScreenShot()

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

''----------------------Tcode MIRO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME23N_0120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_0120_OKCD)

Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_ME23N_1000_COMPANY_CODE)
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
wait(2)
Call SetCombo("RM08M-VORGANG", "Invoice")
Call PressEnter()
'Enter the Posting Date
Call SetTextboxNolabel("INVFO-BLDAT","",ConvertDate(DT_ME23N_0010_INVOICE_DATE),False)
Call PressEnter()
''''Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_ME23N_0010_POSTING_DATE),False)
Call SetTextbox("Reference","INVFO-XBLNR","",DT_ME23N_0010_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_ME23N_0010_REFERENCE,False)
Call SetCombo("RM08M-REFERENZBELEGTYP","Purchase Order/Scheduling Agreement")
Call SetCombo("RM08M-XWARE_BNK","Goods/service items")
Call SetTextboxNoLabel("RM08M-EBELN",0,DT_ME23N_6211_RM08MEBELN,False)
Call SetTextbox("Incg Doc. Nmbr","INVFO-INWARDNO_HD","",DT_ME23N_0010_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()

Call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",DT_ME23N_6310_TABLECELL_AMOUNT_0,False)
Call SetTableData("SAPLMR1MTC_MR1M","Quantity","1","","",DT_ME23N_6310_TABLECELL_QUANTITY_0,False)
'Call TakeScreenShot()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_ME23N_0010_CALCULATE_TAX,False)
Wait(1)
Call PressEnter()
Call TakeScreenShot()

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_BALANCE_OUTPUT",False)
Wait(1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_ME23N_0010_AMOUNT,False)
Call PressEnter()
Wait(1)

'Verify Balance
Call VerifyTextBoxContent("Balance","RM08M-DIFFERENZ",0,DT_ME23N_6000_CHECK_TEXT_OF_BALANCE_OCC1,False)

'Select Row
Call SelectRowGuiTable("SAPLMR1MTC_MR1M","Item",1,False)

'Click on Post Buton
Call SelectMenuBar("Invoice Document;Post")
Call ClickButtonIfExist("Save",True)
wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ME23N_6000_CHECK_TEXT_OF_STATUSBAR)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

