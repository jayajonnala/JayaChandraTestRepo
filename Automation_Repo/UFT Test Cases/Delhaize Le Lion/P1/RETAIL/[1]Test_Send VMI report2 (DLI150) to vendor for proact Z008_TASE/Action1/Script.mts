
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Send VMI report2 (DLI150) to vendor for proact Z008_TASE
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

gstrTestCaseName = "Test_Send VMI report2 (DLI150) to vendor for proact Z008_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

''INPUT''
Call SelectCheckbox("P_SPLIT","0","ON",false)

'Call SetTextbox("Vendor","S_SUPPL-LOW","",DT_ZMDPU_WVMX_1000_VENDOR,false) 
Call SetTextboxNoLabel("S_SUPPL-LOW","",DT_ZMDPU_WVMX_1000_VENDOR,False)
Call SetTextbox("Purch. Organization","P_PURORG","",DT_ZMDPU_WVMX_1000_PURCH_ORGANIZATION,false)  
Call SetTextbox("Site","S_PLANTS-LOW","",DT_ZMDPU_WVMX_1000_SITE,false)
Call SetTextbox("Current Date","P_DATE","",ConvertDate(DT_ZMDPU_WVMX_1000_CURRENT_DATE),false)
Call SetTextbox("Run Number","P_RUN_NO","",DT_ROW_NUMBER,false) 
Call SetTextbox("PROACT control prof.","S_PAPRF-LOW","",DT_ZMDPU_WVMX_1000_PROACT_CONTROL_PROF,false)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()


Call ClickButton("Send IDoc   \(Ctrl\+Shift\+F8\)",False)
Call ClickButton("Continue   \(Enter\)",True)
wait(5)

Call SetTcode(DT_ZMDPU_WVMX_1000_OKCD)     
Call PressEnter()     

Call SetTextbox("Message Variant","MESCOD-LOW","",DT_ZMDPU_WVMX_1100_MESSAGE_VARIANT,false)
Call SetTextbox("Message Function","MESFCT-LOW","",DT_ZMDPU_WVMX_1100_MESSAGE_FUNCTION,false)
Call SetTextbox("Partner Number","PPPRN-LOW","",DT_ZMDPU_WVMX_1100_PARTNER_NUMBER,false)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call TakeScreenShot()
Call SelectColumnGuiGrid("Selected IDocs","",  "IDoc Number", False)
Call TakeScreenShot()
Call ClickButtonToolBar("&SORT_DSC ", "")
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("Selected IDocs", "", 1, False)
Call DoubleClickGuiGridCell("Selected IDocs", "", 1, "IDoc Number", False)
Call TakeScreenShot()

Call ActivateNodeGuiTree(0, "#1;#3;#1")
Call TakeScreenShot()

Call VerifyTextBoxContent("Message","T100-TEXT", "", Lcase(DT_ZMDPU_WVMX_0100_CHECK_TEXT_OF_T100TEXT), False)
Call VerifyTextBoxContent("Status","EDI_INTDS-DESCRP", "", Lcase(DT_ZMDPU_WVMX_0100_CHECK_TEXT_OF_STATUS_OCC1), False)
Call VerifyTextBoxContent("Status","EDI_INTDS-STATUS", "", DT_ZMDPU_WVMX_0100_CHECK_TEXT_OF_STATUS, False)

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

