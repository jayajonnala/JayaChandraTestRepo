

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06DCAOM01_004_DC_Inventory_Management_Destruction_Cancella
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

gstrTestCaseName = "Test_06DCAOM01_004_DC_Inventory_Management_Destruction_Cancella"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------MIGO-----------------------------

call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)
Call TakeScreenShot()
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,Year(Date),False)
wait 2
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0110_POSTING_DATE),False)
Call TakeScreenShot()
Call PressEnter()     ' 

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Where",False)
Call SetTextbox("Reason for Movement","GOITEM-GRUND","",DT_MIGO_0325_REASON_FOR_MOVEMENT,False)

Call TakeScreenShot()
Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_ARTICLE_NO_OUTPUT")
Call VerifyStatusBar("Article document " & DT_ARTICLE_NO_OUTPUT &" posted")

call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC1)
Call TakeScreenShot()
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_ARTICLE_NO_OUTPUT,False)
Call TakeScreenShot()
Call PressEnter()     ' 

Call SelectTab("TS_GOITEM","Where",False)
Call VerifyTextBoxContent("Movement type","GOITEM-BWART",0,DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
Call SelectTab("TS_GOHEAD","Doc. info",False)
Call ClickButton("FI Documents",False)

'''Call SelectRowGuiGridbyRowNo("Documents in Accounting",0,1,True)
'''Call DoubleClickGuiGridCell("Documents in Accounting","",1,"Object type text",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call VerifyGridCellContent("",1,"KTONR",0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"KTONR",0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",1,"LOKKT",0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("",2,"LOKKT",0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)

Call LogOff()
Call FinalStatus ()



