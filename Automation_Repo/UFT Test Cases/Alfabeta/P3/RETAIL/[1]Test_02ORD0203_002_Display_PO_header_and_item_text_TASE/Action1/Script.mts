

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02ORD0203_002_Display_PO_header_and_item_text
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

gstrTestCaseName = "Test_Display_PO_header_and_item_text"
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

''-'---------------------------------  Tcode ME23N ------------------------------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)    
Call PressEnter() 
Call  CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

call ClickButton("Other Purchase Order   \(Shift\+F5\)",False)
Call SetTextbox("Pur. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PO,True)    
Call SelectRadioButton("MEPO_SELECT-BSTYP_F","Pur. Order",True)     
Call ClickButton("Other Document   \(Enter\)",True)     

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
Call SelectTab("HEADER_DETAIL","Texts",False)
Call VerifyTextAreaValue (DT_ME23N_0101_CHECK_TEXT_OF_TEXTEDIT_SHELL)
Call SelectTab("HEADER_DETAIL","Org. Data",False)

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)  
Call SelectTab("ITEM_DETAIL","Texts",False)
Call VerifyTextAreaValue(DT_ME23N_0201_CHECK_TEXT_OF_TEXTEDIT_SHELL)
Call ClickButtonIfExist("Collapse Item Details Ctrl\+F7",False) 


Call LogOff()
Call FinalStatus ()


