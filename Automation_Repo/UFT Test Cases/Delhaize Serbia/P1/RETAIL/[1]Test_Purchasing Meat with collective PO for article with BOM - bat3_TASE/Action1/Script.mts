
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for article with BOM - bat3"
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

gstrTestCaseName = "Test_Purchasing Meat with collective PO for article with BOM - bat3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\RETAIL\DT_Purchasing Meat with collective PO for article with BOM - bat3_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
''----------------------Tcode MIGO----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True) 
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot

Call ClickButton("Messages   \(Shift\+F9\)",False)
Call TakeScreenShot
Call VerifyTableCellContent(3, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)
Call VerifyTableCellContent(4, "Output Type", "SAPDV70ATC_NAST3", DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)

Call SetTcode(DT_ME23N_0100_OKCD) 
Call PressEnter()
Call TakeScreenShot
Call SetTextBox("Document Number","S_EBELN-LOW",0,DT_ME23N_1000_DOCUMENT_NUMBER,False)
Call SetTextBox("Purchasing Organization","S_EKORG-LOW",0,DT_ME23N_1000_PURCHASING_ORGANIZATION,False)
Call SetTextBox("Purchasing Group","S_EKGRP-LOW",0,DT_ME23N_1000_PURCHASING_GROUP,False)
Call SetTextBox("Application","P_KAPPL",0,DT_ME23N_1000_APPLICATION,False)
Call SetTextBox("Processing Status","P_VSTAT",0,DT_ME23N_1000_PROCESSING_STATUS,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call VerifyifGuiLabelExistsByRelativeid(DT_ME23N_0120_CHECK_TEXT_OF_ZNEU,"wnd\[0\]/usr/lbl\[6,6\]")

Call SelectCheckBoxNoLabel(0,"ON",False)
Call ClickButton("Trial Printout   \(Shift\+F4\)",False)
Call TakeScreenShot()



Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

