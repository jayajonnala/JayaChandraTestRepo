
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p2
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
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW33 (Dry goods)_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ME23N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'navigate to "Extras;Delivery Output;Header"
Call SelectMenuBar("Extras;Delivery Output;Header")
Call TakeScreenShot()

'Verify the Output Type
Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_VL03N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Document Flow
Call ClickButton("Document Flow   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

'Click on GUI Tree
Call ActivateItemGuiTree(0,"#1;#1;#1","#1")
Wait(2)
Call TakeScreenShot()

'Get the Document No
Call GetGridContentByTitle("TF to stck in trans",0,"Doc.no.",1,"DT_DOCUMENT_NUM_OUTPUT")

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

