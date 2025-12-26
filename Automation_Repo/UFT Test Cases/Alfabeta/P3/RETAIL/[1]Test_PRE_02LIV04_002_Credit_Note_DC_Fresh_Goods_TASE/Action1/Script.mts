

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_02LIV04_002_Credit_Note_DC_Fresh_Goods
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

gstrTestCaseName = "Test_PRE_02LIV04_002_Credit_Note_DC_Fresh_Goods"
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

'--------------------------------------------  ME21N------------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

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

Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)

Call SetTableData("SAPLMEGUITC_1211","OUn","1","","",DT_ME21N_1211_TABLECELL_OUN_0,False)
Call SetTableData("SAPLMEGUITC_1211","OUn","2","","",DT_ME21N_1211_TABLECELL_OUN_1,False)

Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False)
Call PressEnter()

Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False) 

Call SetTableData("SAPLMEGUITC_1211","Returns Item","1","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_0,False) 
Call SetTableData("SAPLMEGUITC_1211","Returns Item","2","","",DT_ME21N_1211_TABLECELL_RETURNS_ITEM_1,False) 

Call PressEnter()
Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True) 

wait 2
Call GetStatusBar("item2","DT_PO_NUM_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number " & DT_PO_NUM_OUTPUT )
' 
'
'''--------------------------------MIGO-----------------------------
'
''
Call SetTcode(DT_ME21N_0100_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
'
call SetCombo("GODYNPRO-ACTION","Goods Receipt")
call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_ME21N_0010_GODEFAULT_TVBWART,False)
Wait 2
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_PO_NUM_OUTPUT,False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call TakeScreenShot()
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE ,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
Call TakeScreenShot()
CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_ME21N_0304_ITEM_OK,False)
'Call SetTableData("SAPLMIGOTV_GOITEM","OK","1","","",DT_ME21N_0304_ITEM_OK,False) 
Call SetTableData("SAPLMIGOTV_GOITEM","OK","2","","",DT_ME21N_0200_TABLECELL_OK_1,False)
Call PressEnter()   

Call TakeScreenShot()
 Call SelectTab("TS_GOITEM","Batch",False)

Call ClickCellTableByRowNo("SAPLMIGOTV_GOITEM","Line",1,False)
Call  SetGridData("", 1, "DOCUBATCH_CHARG", DT_ME21N_0201_GRIDCELL_0_DB_NO, False)
Call PressEnter()   
Call ClickCellTableByRowNo("SAPLMIGOTV_GOITEM","Line",2,False)

Call  SetGridData("", 1, "DOCUBATCH_CHARG", DT_ME21N_0201_GRIDCELL_0_DB_NO_OCC1, False)
Call PressEnter()   


Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 

Call ClickButtonIfExist("Prev\. screen   \(Ctrl\+F8\)",False)  
Call ClickButtonIfExist("Prev\. screen   \(Ctrl\+F8\)",False)  

Call VerifyStatusBar("posted")
'''
Call LogOff()
Call FinalStatus ()


