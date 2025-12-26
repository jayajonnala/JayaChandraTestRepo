'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

 '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Reload DataSheet to updates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Scrapping in DC_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Scrapping in DC_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Scrapping  in DC_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath) 

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
''
''''----------------------Tcode MIGO----------------------------
''Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
''Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC1)
Wait(1)
''Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
Wait(1)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
Call SelectTab("TS_GOITEM","Article",False)
Call SetTextbox("Article","GOITEM-MAKTX","",DT_MIGO_0310_ARTICLE,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Article","GOITEM-MAKTX",False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MIGO_0300_QUANTITY,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Qty in Unit of Entry","GOITEM-ERFMG","",DT_MIGO_0315_QTY_IN_UNIT_OF_ENTRY,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Qty in Unit of Entry","GOITEM-ERFMG",False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MIGO_0300_WHERE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Site","GOITEM-NAME1","",DT_MIGO_0325_SITE,False)
Call SetTextbox("Storage Location","GOITEM-LGOBE","",DT_MIGO_0325_STORAGE_LOCATION,False)
Call SetTextbox("Reason for Movement","GOITEM-GRUND","",DT_MIGO_0325_REASON_FOR_MOVEMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Reason for Movement","GOITEM-GRUND",False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MIGO_0300_QUANTITY_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MIGO_0300_WHERE_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MIGO_0300_ACCOUNT_ASSIGNMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Cost Center","COBL-KOSTL","",DT_MIGO_1001_COST_CENTER,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Cost Center","COBL-KOSTL",False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_MIGO_0001_ARTICLE_DOCNUM_2_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC3)
Call VerifyStatusBarMessageType("S")

'''''''''''''''''''''''''''''''''''''''''''''''''''''iteration display ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC2)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC","",DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR","",DT_YEAR,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD",DT_MIGO_0100_DOC_INFO,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("FI Documents",False) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

''Call GetGridContentByTitle("Documents in Accounting","","Doc. Number",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER,"DT_MIGO_0200_DOCUMENT_NUMBER_OUTPUT")
Call GetGridContentByRefColumn("Documents in Accounting","","Object type text","Accounting document","Document Number","DT_MIGO_0200_DOCUMENT_NUMBER_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectRowGuiGridbyRowNo("Documents in Accounting","",DT_MIGO_0200_GRIDCELL_0_DOC_NUMBER,true)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",True)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Document Number","BKPF-BELNR",False)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


