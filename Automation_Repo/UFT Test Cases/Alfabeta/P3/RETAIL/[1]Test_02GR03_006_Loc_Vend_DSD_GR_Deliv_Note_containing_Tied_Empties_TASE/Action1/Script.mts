

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
'.................Test Script Name : Test_02GR03_006_Loc_Vend_DSD_GR_Deliv_Note_containing_Tied_Empties
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_02GR03_006_Loc_Vend_DSD_GR_Deliv_Note_containing_Tied_Empties"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR03_006_Loc_Vend_DSD_GR_Deliv_Note_containing_Tied_Empties.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME21N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the Details
Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False) 
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call PressEnter() 
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()

'Enter Purchase Order Details with 3 items'
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_0),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","2","","",DT_ME21N_1211_TABLECELL_ARTICLE_1,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","2","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_1,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","2","","",DT_ME21N_1211_TABLECELL_SITE_1,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","2","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_1),False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Article","3","","",DT_ME21N_1211_TABLECELL_ARTICLE_2,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","3","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_2,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","3","","",DT_ME21N_1211_TABLECELL_SITE_2,False) 
Call PressEnter()
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","3","","",ConvertDate(DT_ME21N_1211_TABLECELL_DELIV_DATE_2),False) 
Call PressEnter()


'Capture the screenshot
Call TakeScreenShot()

'Click on Save Buton
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
VerifyStatusBar("Direct Delivery created under the number " & DT_PO_NUMBER_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''----------------------Tcode MIGO----------------------------
'
Call SetTcode(DT_ME21N_14_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)
Wait(2)
Call SetCombo("GODYNPRO-ACTION","Display")
 ' SetComboByKey(attachedTextOrComboName, keyValue)
Call SetComboByKey("GODYNPRO-REFDOC","R02")
Wait(2)
Call TakeScreenShot()

Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME21N_2000_GODYNPROPO_NUMBER,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MOVEMENT_TYPE,False)
Wait(2)

Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME21N_0110_DELIVERY_NOTE,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
'Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

''------------------------------Display Article Document Details-------------------------------------------------
Call SetCombo("GODYNPRO-ACTION","Display")
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Verify If FB03 Screen is displayed
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)
'Call VerifyGridCellContentbyID("CTRL_CONTAINERBSEG",1,"Assignment",1,DT_ME21N_750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZUONR)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
'***********************************************************************************


