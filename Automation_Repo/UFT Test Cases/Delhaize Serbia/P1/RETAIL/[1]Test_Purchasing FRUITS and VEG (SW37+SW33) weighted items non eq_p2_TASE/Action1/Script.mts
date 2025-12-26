
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p2
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
'.................Test Script Name :Test_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :4th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Purchasing FRUITS and VEG (SW37+SW33) weighted items non eq_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode MIGO----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter the PO Number and Press Enter
Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call TakeScreenShot()
Call PressEnter() 

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call ClickButtonIfexist("Open header data",False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0110_POSTING_DATE),False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call TakeScreenShot()
Call PressEnter() 

Call VerifyCheckBoxValue("GODYNPRO-DETAIL_TAKE","ON")

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")


''----------------------Tcode MIRO----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIGO_0750_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0750_OKCD)

Call SetCombo("RM08M-VORGANG","Invoice")
Wait(2)


'Enter the Posting Date
Call SetTextboxNoLabel("INVFO-BLDAT","",ConvertDate(DT_MIGO_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","INVFO-BUDAT","",ConvertDate(DT_MIGO_0010_POSTING_DATE),False)
Call PressEnter()
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIGO_0010_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Incg Doc\. Nmbr","INVFO-INWARDNO_HD","",DT_MIGO_0010_REFERENCE,False)
Call TakeScreenShot()
Call PressEnter()

'Enter the delivery Note No
Call SetCombo("RM08M-REFERENZBELEGTYP","Delivery Note")
Call SetTextboxNoLabel("RM08M-LFSNR",0,DT_MIGO_6212_RM08MLFSNR,False)
Call TakeScreenShot()
Call PressEnter()

'Select Calculate Tax field as true
Call SelectCheckbox("INVFO-XMWST",0,DT_MIGO_0010_CALCULATE_TAX,False)
Wait(1)
Call PressEnter()
Call TakeScreenShot()

'Get the remaining balance and enter it in Amount Field
Call GetTextboxValue("RM08M-DIFFERENZ",0,"DT_MIRO_BALANCE_OUTPUT",False)
Wait(1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_MIGO_0010_AMOUNT,False)
Call PressEnter()
Wait(1)

'''Verify Balance
Call VerifyTextBoxContent("Balance","RM08M-DIFFERENZ",0,DT_MIGO_6000_CHECK_TEXT_OF_BALANCE_OCC1,False)
'''Call VerifyTextBoxContent("Balance","RM08M-DIFFERENZ",0,"0,00 ",False)
Call SetTableDataNoRef("SAPLMR1MTC_MR1M","#5",1,DT_MIGO_6310_TABLECELL__0,False)

'Click on Post Buton
Call SelectMenuBar("Invoice Document;Post")
Call ClickButtonIfExist("Save",True)
wait(2)

'Validate If Purchase order is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_MIGO_6000_CHECK_TEXT_OF_STATUSBAR)

'----------------------Tcode MIR4----------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIGO_6000_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIGO_6000_OKCD)

'''Display the Invoice Details
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIGO_6150_INVOICE_DOCUMENT_NO,False)
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIGO_6150_FISCAL_YEAR,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Display Document   \(F2\)",False)
wait(2)


'Click on Follow On Document
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(2)

'Verify the screen title
Call CheckScreen(DT_MIGO_6000_OKCD,DT_MIGO_0750_CHECK_TEXT_OF_TITL_OCC1)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)
Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(1)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


