'''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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
'.................Test Script Name : Test_Scrapping in DC_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 10th May
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Scrapping in DC_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Scrapping  in DC_p1_TASE.xls"
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
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
Wait(1)
''Call SetTextbox("TF trfr within plant","GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call SetTextboxNoLabel("GODEFAULT_TV-BWART","",DT_MIGO_0010_GODEFAULT_TVBWART,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE,False)
Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIGO_0390_GOITEMUMLGOBE,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC,False)
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE,False)
Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE,False)
'Capture the screenshot
Call TakeScreenShot()

Call FocusTextBox("Qty in UnE","GODYNPRO-ERFMG",False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTableCellContent(1,"OK","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_SELECTED_OF_TABLECELL_OK_0)

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_MIGO_0001_ARTICLE_DOCNUM_1_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


