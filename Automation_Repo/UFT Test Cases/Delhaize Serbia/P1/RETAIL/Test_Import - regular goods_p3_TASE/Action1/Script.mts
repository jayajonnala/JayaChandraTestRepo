
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Import - regular goods_p3
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
'.................Test Script Name : Test_Import - regular goods_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Import - regular goods_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Import - regular goods_p3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

'.......................Mandatory Initial Call only in First Component in a Test Scenario
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode MIR4----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter() 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Display the Invoice Details
Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIR4_6150_INVOICE_DOCUMENT_NO,False)
'Call SetTextbox("Fiscal Year","RBKP-GJAHR","",DT_MIRO_6150_FISCAL_YEAR,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Display Document   \(F2\)",False)
wait(2)

'Click on Follow On Document
Call ClickButtonIfExist("Follow-On Documents \.\.\.   \(F8\)",False)
wait(2)

'Verify the screen title
Call CheckScreen(DT_MIRO_6000_OKCD,DT_MIR4_0750_CHECK_TEXT_OF_TITL)

'----------------------Tcode MiGO----------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIR4_0750_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIR4_0750_OKCD)

'Enter the PO Number and Press Enter
Call SetCombo("GODYNPRO-ACTION","Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIR4_2000_GODYNPROPO_NUMBER,False)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter the Document Date and Posting date
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIR4_0110_DOCUMENT_DATE),False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIR4_0110_DELIVERY_NOTE,False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIR4_0110_POSTING_DATE),False)
Call TakeScreenShot()
Call PressEnter()

Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_MIR4_0200_CHECK_SELECTED_OF_TABLECELL_OK_0,False)
Call TakeScreenShot()
Call PressEnter()

'Verify checkbox value
Call VerifyCheckBoxValue("GODYNPRO-DETAIL_TAKE",DT_MIR4_0200_CHECK_SELECTED_OF_TABLECELL_OK_0)


'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT&" "&"posted")


''------------------------------Display Article Document Details-------------------------------------------------

Call SetCombo("GODYNPRO-ACTION","Display")
'Call SetCombo("GODYNPRO-REFDOC","Article Document")
'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_ARTICLE_NO_OUTPUT,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,2020,False)
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

'Verify the Title
Call CheckScreen(DT_MIR4_0750_OKCD,DT_MIR4_0750_CHECK_TEXT_OF_TITL_OCC1)


'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)



'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

