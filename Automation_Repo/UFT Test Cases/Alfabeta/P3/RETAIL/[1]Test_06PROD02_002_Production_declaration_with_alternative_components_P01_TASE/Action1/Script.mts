

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06PROD02_002_Production_declaration_with_alternative_components_P01
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
gstrTestCaseName = "Test_06PROD02_002_alternative_components_P01"
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

''--------------------------------MFBF90-----------------------------
 Call SetTextbox("Conf\. Qty","RM61B-ERFMG","",DT_MFBF_0811_CONF_QTY,False)
'Call TakeScreenShot()
' Call SetTextbox("Production Version","RM61B-VERID","",DT_MFBF_0801_PRODUCTION_VERSION,False)
Call TakeScreenShot()
 Call SetTextbox("Article","RM61B-MATNR","",DT_MFBF_0801_ARTICLE,False)
 Call SetTextbox("Site","RM61B-WERKS","",DT_MFBF_0801_SITE,False)
Call TakeScreenShot()
Call PressEnter()     ' 

Call SelectRowGuiTable("SAPLBARMTC160","Ver.",DT_MFBF_0540_TABLECELL_STOR_LOCATION_6,True)
Call ClickButton("Continue   \(Enter\)",False)
Call PressEnter()     ' 

 Call SetTextbox("Docu. Batch No.","RM61B-DOCUBATCH_CHARG","",DT_MFBF_0801_DOCU_BATCH_NO,False)
Call PressEnter()     ' 
 Call TakeScreenShot()
 
Call ClickButton("Post with Correction of Components and/or Activities   \(F8\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_DOC_GI_GR_OUTPUT")
Call VerifyStatusBar("GR and GI with document "& DT_DOC_GI_GR_OUTPUT &" posted")

''--------------------------------MB90-----------------------------
Call SetTcode(DT_MFBF_0800_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MFBF_0800_OKCD)

 Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_DOC_GI_GR_OUTPUT,False)
 Call SetTextbox("Article Doc. Year","PM_MJAHR","",Year(date),False)
Call TakeScreenShot() 
 Call ClickButton("Execute   \(F8\)",False)
 Call VerifyStatusBarMessageType("S")

Call SelectCheckboxNoLabel(0,DT_MFBF_0120_NO_NAME,False)
Call TakeScreenShot()
Call SelectMenuBar("Goto;Print preview")
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()


