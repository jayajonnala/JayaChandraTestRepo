
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Regular purchasing in SW31 dry goods_p2
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
'.................Test Script Name : Test_Regular purchasing in SW31 dry goods_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Regular purchasing in SW31 dry goods_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Regular purchasing in SW31 dry goods_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME23N----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     
Call ClickButtonIfExist("Cancel   (F12)",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Click on Other purchase order
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call TakeScreenShot()

Call ClickButtonIfExist("Other Document   \(Enter\)",True)
wait(2)
Call TakeScreenShot()

'Expand the Header
Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(2)
Call TakeScreenShot()

'Click on Messages button and get the output type.
Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
wait(2)


'Call VerifyTableCellContent(3,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_2)
'Call VerifyTableCellContent(4,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_3)
Call VerifyTableCellContent(4,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_2)
Call VerifyTableCellContent(5,"Status","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_STATUS_3)


'''Verify the Output Type data
'Call VerifyTableCellContent(3,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)
'Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)
Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)
Call VerifyTableCellContent(5,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)


'
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

Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

' SelectCheckboxNoLabel(checkBoxIndex, OnOffStatus, blnIsItPopup)
Call SelectCheckboxNoLabel("0","ON",False)
Call ClickButton("Display Message   \(Shift\+F8\)",False)
Call ClickButton("Back   \(F3\)",False)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
''----------------------Tcode MIGO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_ME23N_0120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_0120_OKCD)

'Enter the PO Number and Press Enter
Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ME23N_2000_GODYNPROPO_NUMBER,False)
Call TakeScreenShot()
Wait(1)
Call PressEnter()
Wait(1)
Call TakeScreenShot()

'Enter the Document details
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_ME23N_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_ME23N_0110_POSTING_DATE),False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_ME23N_0110_DELIVERY_NOTE,False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()
Call PressEnter()

'Verify checkbox value
Call VerifyCheckBoxValue("GODYNPRO-DETAIL_TAKE","ON")

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Wait(2)

Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")

''------------------------------Display Article Document Details-------------------------------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetCombo("GODYNPRO-ACTION","Display")
Call PressEnter() 
Wait(2)
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC","",DT_ME23N_2010_GODYNPROMAT_DOC,False)
Call PressEnter()

Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Verify the Title
Call CheckScreen(DT_ME23N_0120_OKCD,"Display Document: Data Entry View")

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

