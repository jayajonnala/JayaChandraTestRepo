
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Import - regular goods_p4
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
'.................Test Script Name : Test_Import - regular goods_p4
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Import - regular goods_p4"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Import - regular goods_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME23N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,True)
Call PressEnter()     ' 
Call TakeScreenShot
Call SetCombo("RM08M-VORGANG","Invoice")
Call PressEnter()
Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_MIRO_0010_INVOICE_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_0010_REFERENCE,False)
Call PressEnter()
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_MIRO_0010_REFERENCE,False)
Call SetTextBoxNoLabel("INVFO-WAERS",0,DT_MIRO_0010_AMOUNT,False)
Call PressEnter()
Call TakeScreenShot
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call PressEnter() 
Call SetComboByKey("RM08M-XWARE_BNK", DT_MIRO_6211_RM08MXWARE_BNK)
Call PressEnter() 
Call TakeScreenShot
Call SetTextBoxNoLabel("RM08M-EBELN",0,DT_MIRO_6211_RM08MEBELN,False)
Call PressEnter()
Call SetTableData("SAPLMR1MTC_MR1M","Tax Code",1,"Purchase order",DT_MIRO_6211_RM08MEBELN,"QG (20% Procur. - domestic suppliers)",False)
Call SetTableData("SAPLMR1MTC_MR1M","#5",1,"Purchase order",DT_MIRO_6211_RM08MEBELN,"ON",False)

Call SelectCheckBox("INVFO-XMWST","0","ON",False)
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_6000_CHECK_TEXT_OF_BALANCE_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIRO_0010_AMOUNT_OCC1,False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call GetStatusBar("item1","DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document no. "&DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" created")

Call SetTcode(DT_MIRO_6000_OKCD)     
Call PressEnter() 
Call TakeScreenShot()
Call SetComboByKey("GODYNPRO-ACTION", DT_MIRO_0010_GODYNPROACTION)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_TRANSFER_P2P,False)
Call PressEnter()
Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIRO_0390_ARTICLE,False)
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIRO_0390_SITE,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIRO_0390_STOR_LOC,False)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIRO_0390_QTY_IN_UNE,False)
'Call SetTextbox("Site Trfr Pstg","GOITEM-UMNAME1","",DT_MIRO_0390_GOITEMUMNAME1,False)
'Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIRO_0390_GOITEMUMLGOBE,False)
Call SetTextboxNoLabel("GOITEM-UMNAME1","",DT_MIRO_0390_GOITEMUMNAME1,False)
Call SetTextboxNoLabel("GOITEM-UMLGOBE","",DT_MIRO_0390_GOITEMUMLGOBE,False)

Call TakeScreenShot()
Call PressEnter()
Call ClickButton("btn\[11\]",False)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_MIRO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Article document "&DT_MIRO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT& " posted")


Call SetComboByKey("GODYNPRO-ACTION", DT_MIRO_0010_GODYNPROACTION_OCC1)
Call TakeScreenShot()
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_MIRO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)
Call TakeScreenShot()
Call PressEnter()

Call TakeScreenShot()

Call PressEnter()

Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call ClickButton("FI Documents",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

'''
''''Click on Other purchase order
'''Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
'''wait(2)
'''Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
'''Call TakeScreenShot()
'''
'''Call ClickButtonIfExist("Other Document   \(Enter\)",True)
'''wait(2)
'''Call TakeScreenShot()
'''
''''Click on Messages button and get the output type.
'''Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
'''wait(2)
'''
''''Verify the Status
'''Call VerifyTableCellContent(3,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_2)
'''Call VerifyTableCellContent(4,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_3)
'''
''''Verify the Output Type data
'''Call VerifyTableCellContent(3,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)
'''Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)
'''
'''''----------------------Tcode ME9F----------------------------
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''
'''Call SetTcode(DT_ME23N_0100_OKCD) 
'''Call PressEnter()     ' 
'''Call CheckTCodeScreen(DT_ME23N_0100_OKCD)
'''
''''Enter the details
'''Call SetTextbox("Document Number","S_EBELN-LOW","",DT_ME23N_1000_DOCUMENT_NUMBER,False)   
'''Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME23N_1000_PURCHASING_ORGANIZATION,False)   
'''Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_ME23N_1000_PURCHASING_GROUP,False)   
'''Call SetTextbox("Application","P_KAPPL","",DT_ME23N_1000_APPLICATION,False)   
'''Call SetTextbox("Processing Status","P_VSTAT","",DT_ME23N_1000_PROCESSING_STATUS,False)  
'''Call TakeScreenShot()
'''
''''Click on Execute Button
'''Call ClickButton("Execute   \(F8\)",False) 
'''Wait(2)
'''Call TakeScreenShot()
'''
'''''----------------------Tcode MIRO----------------------------
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''Call SetTcode(DT_ME23N_0120_OKCD) 
'''Call PressEnter()     ' 
'''Call CheckTCodeScreen(DT_ME23N_0120_OKCD)
'''
'''Call SetTextboxPopupIfExist("BKPF-BUKRS","Company Code",DT_ME23N_1000_COMPANY_CODE)
'''Call TakeScreenShot()
'''Call ClickButtonIfExist("Continue   \(Enter\)",True)
'''wait(2)
'''
''''Enter the Posting Date
'''Call SetTextbox("Invoice date","INVFO-BLDAT","",ConvertDate(DT_ME23N_0010_INVOICE_DATE),False)
'''Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_ME23N_0010_POSTING_DATE),False)
'''Call SetTextbox("Reference","INVFO-XBLNR","",DT_ME23N_0010_REFERENCE,False)
'''Call TakeScreenShot()
'''Call PressEnter()
'''
''''Enter the delivery Note No
'''Call SetCombo("RM08M-REFERENZBELEGTYP","Purchase Order/Scheduling Agreement")
'''Call SetTextboxNoLabel("RM08M-EBELN",0,DT_ME23N_6211_RM08MEBELN,False)
'''Call TakeScreenShot()
'''Call PressEnter()
'''
'''Call SetTableData("SAPLMR1MTC_MR1M","Amount","1","","",DT_ME23N_6310_TABLECELL_AMOUNT_0,False)
'''Call SetTableData("SAPLMR1MTC_MR1M","Quantity","1","","",DT_ME23N_6310_TABLECELL_QUANTITY_0,False)
''''Call TakeScreenShot()
'''
''''Select Calculate Tax field as true
'''Call SelectCheckbox("INVFO-XMWST",0,DT_ME23N_0010_CALCULATE_TAX,False)
'''Wait(1)
'''Call PressEnter()
'''Call TakeScreenShot()
'''
''''Get the remaining balance and enter it in Amount Field
'''Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_BALANCE_OUTPUT",False)
'''Wait(1)
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''Call SetTextbox("Amount","INVFO-WRBTR","",DT_ME23N_0010_AMOUNT,False)
'''Call PressEnter()
'''Wait(1)
'''
''''Verify Balance
'''Call VerifyTextBoxContent("Balance","RM08M-DIFFERENZ",0,DT_ME23N_6000_CHECK_TEXT_OF_BALANCE_OCC1,False)
'''
''''Select Row
'''Call SelectRowGuiTable("SAPLMR1MTC_MR1M","Item",1,False)
'''
''''Click on Post Buton
'''Call SelectMenuBar("Invoice Document;Post")
'''Call ClickButtonIfExist("Save",True)
'''wait(2)
'''
''''Validate If Purchase order is generated
'''Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'''VerifyStatusBar(DT_ME23N_6000_CHECK_TEXT_OF_STATUSBAR)



'*********************************************End Of Script*********************************************************************

