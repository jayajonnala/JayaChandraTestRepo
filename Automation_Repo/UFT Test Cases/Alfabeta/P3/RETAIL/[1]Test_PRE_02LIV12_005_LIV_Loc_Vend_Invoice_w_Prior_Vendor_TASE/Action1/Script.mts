

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_02LIV12_005_LIV_Loc_Vend_Invoice_w_Prior_Vendor
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

gstrTestCaseName = "Test_PRE_02LIV12_005_LIV_Loc_Vend_Invoice_w_Prior_Vendor"
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
'''
''''''--------------------------------ME21N-----------------------------
'''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False) 
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
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
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False)
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_2),False) 
Call PressEnter()

Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True) 
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC3)
wait 2
Call GetStatusBar("item2","DT_PO_NUM_OUTPUT")
VerifyStatusBar("Direct Delivery created under the number " & DT_PO_NUM_OUTPUT )
 

''--------------------------------MIGO-----------------------------

'
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call SetTcode(DT_ME21N_14_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'
Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUM",Cint(DT_NUM)+1)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

call SetComboByKey("GODYNPRO-ACTION",DT_ME21N_0010_GODYNPROACTION)
call SetComboByKey("GODYNPRO-REFDOC","R01")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULTTV_BWART,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_ME21N_0110_DOCUMENT_DATE),False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK","2","","","ON",False) 
Call SetTableData("SAPLMIGOTV_GOITEM","OK","3","","","ON",False) 
Call PressEnter()   

Call TakeScreenShot()
Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC5)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC5)
Call GetStatusBar("item1","DT_ARTICLE_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ARTICLE_DOCUMENT_OUTPUT &" posted")
'''

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

call SetComboByKey("GODYNPRO-ACTION",DT_ME21N_0010_GODYNPROACTION_OCC2)
Call ClickButton("Execute",False) 
Call PressEnter()   
Call TakeScreenShot()
Call VerifyTextBoxContent("Delivery Note","GOHEAD-LFSNR",0,DT_ME21N_0110_CHECK_TEXT_OF_DELIVERY_NOTE,False)

Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",False) 
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
Call GetTextboxValue("BKPF-BELNR","","DT_ME21N_750_GET_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call VerifyTextBoxContent("Reference","BKPF-XBLNR",0,DT_ME21N_750_CHECK_TEXT_OF_REFERENCE,False)
Call VerifyGridCellContent("",1,"ZUONR",0,DT_ME21N_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)
Call VerifyGridCellContent("",2,"ZUONR",0,DT_ME21N_750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ZUONR)

Call LogOff()
Call FinalStatus ()


