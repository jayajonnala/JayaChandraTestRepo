
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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


gstrTestCaseName = "Test_06-09-67-Manual(ty107)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
gstrInputExcelFilePathAndName="S:\TASETestData\DLL\RETAIL\DT_06-09-67-Manual GR in SAP ( mvt type107)_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


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
'Call SetComboByKey("GODYNPRO-ACTION", DT_ME21N_1105_MEPO_ACTION)
'Call PressEnter() 
'Call SetComboByKey("GODYNPRO-REFDOC", DT_ME21N_1105_MEPO_ACTION_REFDOC)
'
'Call PressEnter()
Call SetCombo("GODYNPRO-ACTION", "Goods Receipt")
Call TakeScreenshot()
Call SetCombo("GODYNPRO-REFDOC", "Purchase Order")
Call PressEnter()
Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2000_GODYNPROPO_NUMBER, False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART", 0,DT_MIGO_0010_GODEFAULT_TVBWART, False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call PressEnter()
Call TakeScreenshot()
Call PressEnter()
Call VerifyTableCellContent(1, "Article","SAPLMIGOTV_GOITEM",DT_MIGO_CHECK_ARTICLE_0)
Call VerifyTableCellContent(2, "Article","SAPLMIGOTV_GOITEM",DT_MIGO_CHECK_ARTICLE_1)
Call VerifyTableCellContent(1, "OK","SAPLMIGOTV_GOITEM","ON")
Call VerifyTableCellContent(2, "OK","SAPLMIGOTV_GOITEM","ON")
Call TakeScreenshot()
Call ClickButton("Check Entries   \(F7\)",false)
'Call VerifyStatusBar(DT_ME21N_0001_CHECK_TEXT_OF_STATUSBAR)

'Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call ClickButton("Post   \(Ctrl\+S\)",false)
Call TakeScreenshot()
Call GetStatusBar("item1","DT_ME21N_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call LogOff()
Call FinalStatus ()
'''control data

