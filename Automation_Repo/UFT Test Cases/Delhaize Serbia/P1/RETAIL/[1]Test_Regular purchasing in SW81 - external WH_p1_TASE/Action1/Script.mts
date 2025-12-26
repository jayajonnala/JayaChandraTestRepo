
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Regular purchasing in SW81 - external WH_p1
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Regular purchasing in SW81 - external WH_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Regular purchasing in SW81 - external WH_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Regular purchasing in SW81 - external WH_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'----------------------Tcode MIGO----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter Transfer Posting Details
Call SetCombo("GODYNPRO-ACTION","Transfer Posting")
Call SetCombo("GODYNPRO-REFDOC","Other")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)
Call PressEnter()
Call TakeScreenShot()

Call ClickButtonIfExist("BUTTON_DETAIL",False) 

Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE,False) 
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE,False) 
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC,False) 
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE,False) 
Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIGO_0390_GOITEMUMLGOBE,False)
Call TakeScreenShot()
Call PressEnter()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_DOC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)


'Enter TRansfer Posting Details
Call SetCombo("GODYNPRO-ACTION","Transfer Posting")
Call SetCombo("GODYNPRO-REFDOC","Other")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART_OCC1,False)
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_0390_ARTICLE_OCC1,False) 
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_0390_SITE_OCC1,False) 
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_0390_STOR_LOC_OCC1,False) 
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_0390_QTY_IN_UNE_OCC1,False) 
'Call SetTextbox("Site Trfr Pstg","GOITEM-UMNAME1","",DT_MIGO_0390_GOITEMUMNAME1,False)
Call SetTextbox("Plant Trfr Pstg","GOITEM-UMNAME1","",DT_MIGO_0390_GOITEMUMNAME1,False)
Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIGO_0390_GOITEMUMLGOBE_OCC1,False)
Call TakeScreenShot()
Call PressEnter()


'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_DOC_NO_OUTPUT_1")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC3)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

