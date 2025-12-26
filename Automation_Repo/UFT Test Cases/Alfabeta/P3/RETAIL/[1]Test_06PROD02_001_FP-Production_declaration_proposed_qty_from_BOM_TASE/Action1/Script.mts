

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06PROD02_001_FP-Production_declaration_proposed_qty_from_BOM
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
gstrTestCaseName = "Test_06PROD02_001_FP_qty_from_BOM"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MFBF-----------------------------
 Call SetTextbox("Conf\. Qty","RM61B-ERFMG","",DT_MFBF_0811_CONF_QTY,False)
Call TakeScreenShot()
 Call SetTextbox("Production Version","RM61B-VERID","",DT_MFBF_0801_PRODUCTION_VERSION,False)
Call TakeScreenShot()
 Call SetTextbox("Article","RM61B-MATNR","",DT_MFBF_0801_ARTICLE,False)
 Call SetTextbox("Site","RM61B-WERKS","",DT_MFBF_0801_SITE,False)
Call TakeScreenShot()
Call PressEnter()     ' 

Call ClickButton("Post with Correction of Components and/or Activities   \(F8\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_ORDER_NUMBER_OUTPUT")
Call VerifyStatusBar("GR and GI with document "& DT_ORDER_NUMBER_OUTPUT &" posted")

''--------------------------------MF12-----------------------------
Call SetTcode(DT_EXPECTEDTRANSACTIONCODE_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
'
 Call SetTextbox("Article","SO_MATNR-LOW","",DT_MFBF_1000_ARTICLE,False)
 Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_MFBF_1000_POSTING_DATE),False)
Call PressEnter()     ' 
Call TakeScreenShot() 
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot() 
'Call SetFocusGuiLabel(DT_ORDER_NUMBER_OUTPUT,"","",False)
Call SetFocusGuiLabel("",144,120,False)

Call SendKey("{F2}")
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",False) 
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("Documents in Accounting",0,1,True)
Call DoubleClickGuiGridCell("Documents in Accounting","",1,"Object type text",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call LogOff()
Call FinalStatus ()


