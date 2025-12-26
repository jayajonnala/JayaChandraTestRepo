'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_02-06-01-15-Manual price calculation_TASE
'.................Test Scenario: AT_Sales price with exception rule_TASE
'.................TCode: VKP5
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "TC2_Test_02-06-01-15-Manual"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.

'//----------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------//

'Call CloseSessionsSAP()
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)'.......................Mandatory Initial Call only in First Component in a Test Scenario

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'TCode
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("Article","S_MATNR-LOW","",DT_VKP5_1000_ARTICLE,False)
Call SetTextbox("Sales organization","S_VKORG-LOW","",DT_VKP5_1000_SALES_ORGANIZATION,False)
Call SetTextbox("Distribution channel","S_VTWEG-LOW","",DT_VKP5_1000_DISTRIBUTION_CHANNEL,False)
Call SetTextbox("Validity","S_DATUM-LOW","",ConvertDateFormat(DT_VKP5_1000_VALIDITY),False)
Call SetTextbox("to","S_DATUM-HIGH","",ConvertDateFormat(DT_VKP5_1000_TO),False)
Call TakeScreenShot
Call SetTextbox("Purchase Price Determ\. Seq\.","P_EKERV","",DT_VKP5_1000_PURCHASE_PRICE_DETERM_SEQ,False)
Call PressEnter()
Call ClickButton("Execute   \(F8\)",False)
Call VerifyTextBoxContent("Article","","3",DT_VKP5_0120_CHECK_TEXT_OF_NO_NAME_OCC1,False)
Call VerifyifGuiLabelExistsByRelativeid(DT_VKP5_0120_CHECK_TEXT_OF_BS10,"wnd\[0\]/usr/lbl\[3,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_VKP5_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[8,8\]")
Call SelectCheckboxNoLabel(0,"ON",False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Enter   \(F5\)",True)    
Call ClickButtonIfExist("Enter   \(F5\)",True) 
Call ClickButtonIfExist("Enter   \(F5\)",True) 
Call TakeScreenShot
Call GetStatusbar("item1","DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_VKP5_1000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call TakeScreenShot
Call VerifyStatusBar(DT_VKP5_1000_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus()


