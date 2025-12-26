

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV01_013_LocVend_DC-DSD_Invoice_Free_Goods_w_value
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
	GetRowNo= Parameter("datatable_row")	
End If


If qtpParamExist("RunTimeResultFolder") Then
    RunTimeResultFolder= Parameter("RunTimeResultFolder")    
End If

gstrTestCaseName = "Test_02LIV01_013_LocVend_DC-DSD_Invoice_Free_Goods_w_value"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------------------  ME21N------------------------------------------------
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)   
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)   

Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False)
Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False) 

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)


Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_2),False) 
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Free","1","","",DT_FREE_ITEM,False) 

Call SetTableData("SAPLMEGUITC_1211","Free","2","","",DT_FREE_ITEM,False) 


Call ClickButton("Check   \(Ctrl\+Shift\+F3\)",False)
Call VerifyStatusBar(DT_ME21N_14_CHECK_TEXT_OF_STATUSBAR)

Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True) 

wait 2
Call GetStatusBar("item2","DT_PO_NUM_OUTPUT")
Call VerifyStatusBar("Direct Delivery created under the number " & DT_PO_NUM_OUTPUT )
 

''--------------------------------MIGO-----------------------------

'
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_ME21N_14_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'
call SetComboByKey("GODYNPRO-ACTION",DT_ME21N_0010_GODYNPROACTION)
'call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_ME21N_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_PO_NUM_OUTPUT,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_ME21N_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_ME21N_0110_POSTING_DATE),False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
DeliveryNote= DT_ME21N_0110_DELIVERY_NOTE
Call ClickButtonIfExist("MIGO_OK_GO",False)
CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)

 
Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyStatusBar(DT_ME21N_1_CHECK_TEXT_OF_STATUSBAR)


Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ART_DOC_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ART_DOC_NO_OUTPUT &" posted")
'''
'''--------------------------------MIRO-----------------------------

'
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'
Call SetTcode(DT_ME21N_1_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_ME21N_1_OKCD)

 Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_ME21N_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetCombo("RM08M-VORGANG","Invoice")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_ME21N_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_ME21N_6020_RM08MREFERENZBELEGTYP)
Call SetTextboxNoLabel("RM08M-LFSNR",0,DeliveryNote,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_ME21N_0010_CALCULATE_TAX,False)
Call PressEnter()

Call VerifyTableCellContent(1,"Amount","SAPLMR1MTC_MR1M",DT_ME21N_6310_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)

Call PressEnter()
Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  

Call GetStatusBar("item1","DT_DOC_NUM_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_DOC_NUM_OUTPUT & " created")

'''''--------------------------------WE09-----------------------------
'
Call SetTcode(DT_ME21N_6000_OKCD_OCC3)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
 Call SetTextbox("Direction \(1=Outb\., 2=Inb\.\)","DIRECT-LOW","",DT_MR8M_1000_DIRECTION__1OUTB__2INB,False)

Call SetTextbox("Basic Type","IDOCTP-LOW","",DT_BASIC_TYPE_LOW,False)
 Call SetTextbox("Logical Message","MESTYP-LOW","",DT_LOGICAL_MESSAGE_LOW,False)
 Call SetTextbox("Search in Segment \.\.\.","SEGMENT1","",DT_SEARCH_IN_SEGMENT,False)
 Call SetTextbox("Search in Field \.\.\.","FIELD1_1","",DT_SEARCH_IN_FIELD,False)
 Call SetTextbox("to","STATUS-HIGH","",DT_STATUS,False) 
  Call SetTextbox("for Value \.\.\.","VALUE1_1",0,DT_DOC_NUM_OUTPUT,False)
  
 Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call VerifyStatusBar("IDocs were found")
Call VerifyifGuiLabelExists(DT_ME21N_120_CHECK_TEXT_OF_DATA_PASSED_TO_PORT_OK)
Call GetLabelContentByRefLabel(DT_LOGICAL_MESSAGE_LOW,994,0,"DT_IDOC_OUTPUT",False)
Call ClickLabel(DT_IDOC_OUTPUT,0,False)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RBKP","ZZMDIV_E1RBKP","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","BELNR","Fld Cont.",DT_DOC_NUM_OUTPUT)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;ZZMDIV_E1RSEG","ZZMDIV_E1RSEG","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","EBELN","Fld Cont.",DT_PO_NUM_OUTPUT)

Call ClickLinkGuiTree("IDoc "& DT_IDOC_OUTPUT&";Data records;#2;ZZMDIV_E1MBGM","ZZMDIV_E1MBGM","",False)
Call VerifyTableCellContentByRef("IDOC_TREE_CONTROLINT_SEG_CONTROL","Fld Name","MBLNR","Fld Cont.",DT_ART_DOC_NO_OUTPUT)


'''''--------------------------------MIR4-----------------------------
Call SetTcode(DT_ME21N_100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_DOC_NUM_OUTPUT,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_ME21N_6150_FISCAL_YEAR,False)
Call ClickButton("Display Document   \(F2\)",False)
Call ClickButton("Follow-On Documents \.\.\.   \(F8\)",False) 

'Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Doc. Number","DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT")
Call TakeScreenShot()
Call  ClickButton("Cancel   \(F12\)",True)

'''''--------------------------------J1GP01-----------------------------

Call SetTcode(DT_ME21N_6000_OKCD_OCC4)  
Call PressEnter() 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC7)

Call SetTextbox("Document","S_BELNR-LOW","",DT_MIRO_200_GETCELLVALUE_DOCNR_OUTPUT,False)
Call SetTextbox("Company Code","S_BUKRS","",DT_ME21N_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","S_GJAHR","",CSTR(Year(Date)),False)
Call SelectCheckbox("S_TEST",0,DT_ME21N_1000_TEST_MODE,False)
Call  ClickButton("Execute   \(F8\)",False)
Call VerifyStatusBar("posted")
Call GetStatusBar("item1","DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT")
Call VerifyifGuiLabelExists( "Document "&DT_MIRO_120_GETTEXT_PU_NBR_OUTPUT&" was posted in company code "& DT_ME21N_1000_COMPANY_CODE)

Call LogOff()
Call FinalStatus ()





