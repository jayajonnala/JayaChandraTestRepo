

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_02LIV07_001_LIV_Loc_Vend_DC_Cancellation_Dry_Goods
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


gstrTestCaseName = "Test_PRE_02LIV07_001_Dry_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

'''''--------------------------------ME21N-----------------------------

'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)



Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False) 
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_ME21N_1211_TABLECELL_OUN_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
 Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call PressEnter()
Call TakeScreenShot()

Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)    

wait 2
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
VerifyStatusBar("Standard PO Retail created under the number " & DT_PO_NUMBER_OUTPUT)
	


''--------------------------------MIGO-----------------------------

'
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_ME21N_0100_OKCD_OCC1)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)
'
call SetCombo("GODYNPRO-ACTION","Goods Receipt")
call SetCombo("GODYNPRO-REFDOC","Purchase Order")
'Call SetTextbox("GR goods receipt","GODEFAULT_TV-BWART","",DT_ME21N_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_ME21N_0010_GODEFAULT_TVBWART,False)

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_ME21N_0304_ITEM_OK,False)
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBar("posted")
'''
'''--------------------------------MIRO-----------------------------

'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'
Call SetTcode(DT_ME21N_0100_OKCD_OCC2)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC6)

 Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_ME21N_0010_INVOICE_DATE),False) '------ Datelabel changes as per transaction hence attached text not used
Call SetCombo("RM08M-VORGANG","Invoice")
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_ME21N_0010_REFERENCE,False)
Call PressEnter()
Call SetComboByKey("RM08M-REFERENZBELEGTYP", DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call PressEnter()
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_ME21N_6212_RM08MLFSNR,False)
Call SelectCheckbox("INVFO-XMWST",0,DT_ME21N_0010_CALCULATE_TAX,False)
Call GetTextboxValue("RM08M-DIFFERENZ","","DT_GET_AMOUNT_OUTPUT",False)

'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call PressEnter()
Call SetTextbox("Amount","INVFO-WRBTR","",ConvertNegativePosetive(DT_GET_AMOUNT_OUTPUT),False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_INVOICE_OUTPUT")
Call VerifyStatusBar("Document no. "& DT_INVOICE_OUTPUT & " created")

Call LogOff()
Call FinalStatus ()




