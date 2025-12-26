
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Wholesale with WMS delivery - Integration_p2
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
'.................Test Script Name :Test_Wholesale with WMS delivery - Integration_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Wholesale with WMS delivery - Integration_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Wholesale with WMS delivery - Integration_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode VL03N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'navigate to "Extras;Delivery Output;Header"
Call SelectMenuBar("Extras;Delivery Output;Header")
Call TakeScreenShot()


Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_VL03N_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_VL03N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(2)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(2)

''------------------------------TCode vl02n-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_VL03N_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_VL03N_0100_OKCD)


'Enter the delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY_OCC1,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)

Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Wait(2)

Call GetStatusBar("item3","DT_ARTICLE_DOCUMENT_NUM_OUTPUT")
Wait(2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_VL03N_4004_OUTBOUND_DELIVERY_OCC1&" saved, article document "&DT_VL02N_0200_GRIDCELL_0_ART_DOC_NUMBER&" created")


'Call VerifyStatusBar(DT_VL03N_4004_CHECK_TEXT_OF_STATUSBAR)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Wait(2)


''------------------------------TCode vl03n-------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_VL03N_0100_OKCD_OCC1) 
Call PressEnter() 
Call CheckTCodeScreen(DT_VL03N_0100_OKCD_OCC1)


'Enter the delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY_OCC2,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()

'Click on Document Flow
Call ClickButton("Document Flow   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

'Click on GUI Tree
Call ActivateItemGuiTree(0,"#1;#1;#1","#1")
Wait(2)
Call TakeScreenShot()

'Get the Document No
Call GetGridContentByTitle("GD goods issue:delvy.*",0,"Doc.no.",1,"DT_DOCUMENT_NUM_OUTPUT")


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


