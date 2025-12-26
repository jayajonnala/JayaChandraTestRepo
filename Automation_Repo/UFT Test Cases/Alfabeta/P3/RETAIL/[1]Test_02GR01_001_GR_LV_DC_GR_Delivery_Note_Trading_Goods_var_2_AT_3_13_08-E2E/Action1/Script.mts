

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02GR01_001_GR_LV_DC_GR_Delivery_Note_Trading_Goods_var_2
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

gstrTestCaseName = "Test_02GR01_001_GR_Delivery_Note_2"
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

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MIGO-----------------------------


'
call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call SetTextbox("HeaderText","GOHEAD-BKTXT","",DT_MIGO_0110_HEADERTEXT,False)
Call ClickButtonIfExist("MIGO_OK_GO",False)

Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)

If DT_Multi_Article1<>"" Then
	Call SetTableData("SAPLMIGOTV_GOITEM","OK","2","","",DT_Multi_Article1,False) 
	
End If

If DT_Multi_Article2<>"" Then
	Call SetTableData("SAPLMIGOTV_GOITEM","OK","3","","",DT_Multi_Article2,False) 
	
End If
Call PressEnter()     ' 
Call SelectTab("TS_GOITEM","Where",False)
Call ClickButtonIfExist("Check Entries   \(F7\)",False)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

'Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call VerifyTextBoxContent("Movement type","GOITEM-BWART","",DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
Call VerifyTextBoxContent("Site","GOITEM-WERKS","",DT_MIGO_0325_CHECK_TEXT_OF_SITE,False)

'Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_ARTICLE_DOC_OUTPUT")
Call VerifyStatusBar("Article document "& DT_ARTICLE_DOC_OUTPUT &" posted")


Call LogOff()
Call FinalStatus ()


