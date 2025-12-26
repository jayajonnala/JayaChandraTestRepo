

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02ORD0202_007_Change_item_or_header_text
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

gstrTestCaseName = "Test_02ORD0202_007_Change_item_or_header_text"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

''---------------------------------  Tcode ME22N ------------------------------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)    
Call PressEnter() 
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE,"")

Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUM",Cint(DT_NUM)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButtonIfExist("Switch Off Document Overview   \(F9\",False)

call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME22N_0003_PO,True)    
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
Call ClickButton("Other Document   \(Enter\)",True)   
Call ClickButtonIfExist("No",True)
Call TakeScreenShot()


Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call GetspecialTextboxValue("MEPO_TOPLINE-SUPERFIELD","","DT_ME22N_1105_GET_TEXT_OF_VENDOR_OUTPUT",False)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
If (VerifyTextBoxEnabled("Purch\. Org\.","MEPO1222-EKORG","",False)=False) Then
Call ClickButtonIfExist("Display/Change   \(F7\)",False)
End If
Call TakeScreenShot()
Call GetTextboxValue("MEPO1222-EKORG","","DT_ME22N_1221_GET_TEXT_OF_PURCH_ORG_OUTPUT",False)
Call GetTextboxValue("MEPO1222-EKGRP","","DT_ME22N_GET_CHECK_TEXT_OF_PURCH_GROUP_OUTPUT",False)
Call GetTextboxValue("MEPO1222-BUKRS","","DT_ME22N_1221_CGET_TEXT_OF_COMPANY_CODE_OUTPUT",False)
Call GetTextboxValue("MEPO_TOPLINE-BEDAT","","DT_ME22N_1221_CGET_TEXT_OF_DOC_DATE_OUTPUT",False)
DOC_DATE_OUTPUT=DT_ME22N_1221_CGET_TEXT_OF_DOC_DATE_OUTPUT
PURCH_GROUP_OUTPUT=DT_ME22N_GET_CHECK_TEXT_OF_PURCH_GROUP_OUTPUT
Call TakeScreenShot()
Call GetTableCellData("SAPLMEGUITC_1211","Article",1,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_ARTICLE_0_OUTPUT",False)
Call GetTableCellData("SAPLMEGUITC_1211","Article",2,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_ARTICLE_1_OUTPUT",False)
Call GetTableCellData("SAPLMEGUITC_1211","Article",3,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_ARTICLE_2_OUTPUT",False)
Call GetTableCellData("SAPLMEGUITC_1211","Deliv. Date",1,"","","DT_ME22N_1211_GET_TEXT_OF_TABLECELL_DELIV_DATE_0_OUTPUT",False)
DELIV_DATE_0_OUTPUT=DT_ME22N_1211_GET_TEXT_OF_TABLECELL_DELIV_DATE_0_OUTPUT
Call TakeScreenShot()
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call SelectTab("HEADER_DETAIL","Texts",False)    
Call SelectNodeGuiTree("","Header text")     
Call SetTextArea(DT_ME22N_0101_TEXTEDIT_SHELL)  
Call SelectTab("HEADER_DETAIL","Org. Data",False)    
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Texts",False)     
Call SelectNodeGuiTree("1","Item text")     
Call SetTextArea(DT_ME22N_0201_TEXTEDIT_SHELL)     
Call SelectTab("ITEM_DETAIL","Material Data",False) 
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call TakeScreenShot()
Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)  
Call TakeScreenShot()
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call VerifyTextAreaValue(DT_ME22N_0101_CHECK_TEXT_OF_TEXTEDIT_SHELL)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Texts",False)
Call VerifyTextAreaValue(DT_ME22N_0201_CHECK_TEXT_OF_TEXTEDIT_SHELL)
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  

'---------------------------------  Tcode WWp1 ------------------------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME22N_14_OKCD)    
Call PressEnter() 
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2,"")
Call SelectMenuBar("Goto;Variants;Get...")
Call TakeScreenShot()
Call SetTextbox("Created By","ENAME-LOW","","",True) 
Call SetTextbox("Original Language","MLANGU-LOW","","",True) 
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)     ' - Line (9)
Call SelectRowGuiGrid("Variant Catalog for Program RWWP001.*","","Variant name","AB",True)                  'Please provide reference column name and value     ' - Line (10)
Call PressEnter() 
Call TakeScreenShot()
Call SetTextbox("Site","01SWERKS-LOW","",DT_ME22N_0001_SITE,False)     
'Call SetTextbox("Vendor","01SLIFNR-LOW","",DT_ME22N_0001_VENDOR,False) 
Call SetTextboxNoLabel("01SLIFNR-LOW","",DT_ME22N_0001_VENDOR,False)
Call SetTextbox("Purchasing Group","01SSELEG-LOW","",PURCH_GROUP_OUTPUT,False)    
 Call TakeScreenShot()   
Call SetTextbox("Delivery date","01PEINDT","",ConvertDate(DELIV_DATE_0_OUTPUT),False)  
Call SetTextbox("Purchase Order Date","01PBEDAT","",ConvertDate(DOC_DATE_OUTPUT),False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 
Call PressEnter() 
Call ClickButtonToolBar("SEA","")    
Call SetTextbox("Find:","G_STRING","",DT_ME22N_102_FIND,True)  
Call ClickButton("OK   \(Enter\)",True)   
Call TakeScreenShot()
Call ClickButtonToolBar("EXE","")  
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2,"")
Call TakeScreenShot()
Call ClickButton("Additional Functions   \(Shift\+F9\)",False)  
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE,"")
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call SelectTab("HEADER_DETAIL","Texts",False)    
Call SelectNodeGuiTree("","Header text")     
Call SetTextArea(DT_ME22N_0101_TEXTEDIT_SHELL_OCC2)   
Call TakeScreenShot()
Call SelectTab("HEADER_DETAIL","Org. Data",False)     
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Texts",False)    
Call SelectNodeGuiTree("1","Item text")     
Call SetTextArea(DT_ME22N_0201_TEXTEDIT_SHELL_OCC2)    
Call SelectTab("ITEM_DETAIL","Material Data",False) 
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call TakeScreenShot()

Call SelectMenuBar("Purchase Order;Save")
Call ClickButtonIfExist("Save",True)
Call  CheckScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4,"")
Call SetTcode(DT_SAPTRANSACTIONCODE)      
Call PressEnter() 
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call VerifyTextAreaValue(DT_ME22N_0101_CHECK_TEXT_OF_TEXTEDIT_SHELL_OCC2)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot()
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Texts",False)
Call VerifyTextAreaValue(DT_ME22N_0201_CHECK_TEXT_OF_TEXTEDIT_SHELL_OCC2)
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False)  
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


