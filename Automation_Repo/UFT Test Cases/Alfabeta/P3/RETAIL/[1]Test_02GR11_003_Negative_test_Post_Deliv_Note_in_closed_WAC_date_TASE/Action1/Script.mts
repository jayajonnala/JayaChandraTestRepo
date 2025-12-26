

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02GR11_003_Negative_test_Post_Deliv_Note_in_closed_WAC_date
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

gstrTestCaseName = "Test_02GR11_003_Negative_test_Post_Deliv_Note_in_closed_WAC_date"
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
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'--------------------------------SM34-----------------------------
Call TakeScreenshot()
Call SetTextbox("View cluster","VCLDIR-VCLNAME","",DT_SM34_200_VIEW_CLUSTER,False)
Call TakeScreenshot()
Call ClickButton("Display",False)
wait 3
Call TakeScreenshot()
Call SelectRowGuiTable("SAPLZACPF_TAB_MNTTCTRL_ZACPF_V_LOG_LOGV","Logical Group","ABE002",False)
Call TakeScreenshot()
Call ActivateItemGuiTree(0,"Groups & LOG options;Parameters","Parameters")
Call TakeScreenshot()
Call GetTableCellData("SAPLZACPF_TAB_MNTTCTRL_ZACPF_V_LG_DV_PV","Parameter Value",1,"","","DT_SM34_130_GET_WAC_DATE_OUTPUT",False)
Call TakeScreenshot()
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)

'--------------------------------MIGO-----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_SM34_100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_SM34_0010_GODEFAULT_TVBWART_OCC2,False)
Call TakeScreenshot()
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_SM34_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_SM34_0110_DELIVERY_NOTE,False)
Call TakeScreenshot()
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_SM34_0110_DOCUMENT_DATE_OCC4),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_SM34_0110_POSTING_DATE_OCC4),False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
Call TakeScreenshot()
Call SetTableData("SAPLMIGOTV_GOITEM","OK",1,"","",DT_SM34_0304_ITEM_OK_OCC3,False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_SM34_0304_ITEM_OK_OCC5,False)
Call TakeScreenshot()
Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyifGuiLabelExists(DT_CHECK_ERROR_MESSAGE_CLOSED_DATE)
Call TakeScreenshot()
Call ClickButton("Continue   \(Enter\)",False)


Call LogOff()
Call FinalStatus ()

