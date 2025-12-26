

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_MD_V_AB_001_ Manage Vendor Lifecycle-Create Vendor (Direct Local no Subrange_Retail)_P02_MDM_CLNT_EXTR
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


gstrTestCaseName = "Test_MD_V_AB_001_P02_MDM_CLNT_EXTR"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.04.02.21 VIM - PO Precontrole Issue - BR01 - Invalid Company_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =3
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  MDM_CLNT_EXTR------------------------------------------------
Call SetTextbox("Variant","MDMEXVARI_TS-VARIANTE","",DT_MDM_CLNT_EXTR_0100_VARIANT,False)
Call SetTextbox("Description","MDMEXVARI_TS-REQ_TXT","",DT_MDM_CLNT_EXTR_0100_DESCRIPTION,False)
Call SetTextbox("Extraction Object","MDMEXVARI_TS-EXTRACTOR","",DT_MDM_CLNT_EXTR_0100_EXTRACTION_OBJECT,False)
Call SetTextbox("Target System","MDMEXVARI_TS-MESSENGERID","",DT_MDM_CLNT_EXTR_0100_TARGET_SYSTEM,False)
Call SetTextbox("Distribution Mode","MDMEXVARI_TS-UPDMODE","",DT_MDM_CLNT_EXTR_0100_DISTRIBUTION_MODE,False)
Call SetTextbox("Block Size","MDMEXVARI_TS-BLOCKSIZE","",DT_MDM_CLNT_EXTR_0100_BLOCK_SIZE,False)

Call PressEnter() 
Call ActivateItemGuiTree(0,"#1","#2")
Call SetGridData(gridTitle,3,"From Value",DT_MDM_CLNT_EXTR_0100_GRIDCELL_2_FROM_VALUE,False)

'SAPGuiSession("Session").SAPGuiWindow("Local Control of Extraction").SAPGuiTree("TableTreeControl").ActivateItem "E1LFA1M","#2"
'SAPGuiSession("Session").SAPGuiWindow("Local Control of Extraction").SAPGuiGrid("GridViewCtrl").SelectCell 3,"From Value"

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call ClickButton("Start Extraction   \(F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call VerifyStatusBar(DT_MDM_CLNT_EXTR_0100_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()



