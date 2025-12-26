

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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02GR03_004_Loc_Vend_DSD_GR_Deliv_Note_w_Trading_Goods
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR03_004_Loc_Vend_DSD_GR_Deliv_Note_w_Trading_Goods"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_004_Loc_Vend_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
'DataRowSet=3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME21N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False) 
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)   
Call TakeScreenShot()
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()
Call TakeScreenShot()

'Enter Purchase Order Details with 3 items'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False) 
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False) 
Call PressEnter()
Call TakeScreenShot()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_2),False) 
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Check")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC4)
Call ClickButtonIfExist("Save",True)
wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
Call VerifyStatusBar("Direct Delivery created under the number " & DT_PO_NUMBER_OUTPUT)

'----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME21N_14_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Wait(2)
Call TakeScreenShot()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUM",Cint(DT_NUM)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_PO_NUMBER_OUTPUT,False)
Call TakeScreenShot()
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULTTV_BWART,False)
Wait(2)

Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)

Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

Call ClickButtonIfExist("MIGO_OK_GO",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC5)
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")

''------------------------------Display Article Document Details-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetCombo("GODYNPRO-ACTION","Display")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ARTICLE_DOCUMENT,False)
Call PressEnter() 
Call VerifyTextBoxContent("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Wait(2)
Call TakeScreenShot()

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

''KGARA on 28/12/2022.
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'Call GetTextboxValue("Document Number","",DT_ME21N_750_GET_TEXT_OF_DOCUMENT_NUMBER_OUTPUT,False)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_ME21N_750_CHECK_TEXT_OF_REFERENCE,False)
Call VerifyGridCellContent("",1,"Assignment","",DT_ME21N_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"Assignment","",DT_ME21N_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

