

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02GR11_001_Negative_test_Duplicate_Deliv_Note
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

gstrTestCaseName = "Test_02GR11_001_Negative_test_Duplicate_Deliv_Note"
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

'--------------------------------MIGO-----------------------------

call SetComboByKey("GODYNPRO-ACTION",DT_EXECUTABLE_ACTION)
call SetCombo("GODYNPRO-REFDOC","Purchase Order")
Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0110_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0110_POSTING_DATE),False)
Call ClickButtonIfExist("MIGO_OK_GO",False)
'
Call ClickButtonIfExist("Open detail data",False)

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_TRUE,False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK","1","","",DT_TRUE,False)
Call PressEnter()     ' 

Call ClickButton("Check Entries   \(F7\)",False)
Call VerifyStatusBar(DT_MIGO_1_CHECK_TEXT_OF_STATUSBAR)

Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Call GetStatusBar("item1","DT_ARTICLE_DOCUMENT_OUTPUT")
Call VerifyStatusBar("Article document "& DT_ARTICLE_DOCUMENT_OUTPUT &" posted")

'------------------------------------------------- Migo ------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextboxNoLabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART_OCC2,False)

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER_OCC2,False)
Call SetTextbox("Delivery Note","GOHEAD-LFSNR","",DT_MIGO_0110_DELIVERY_NOTE_OCC2,False)
Call SetTextbox("Document Date","GOHEAD-BLDAT","",ConvertDate(DT_MIGO_0110_DOCUMENT_DATE_OCC2),False)
Call SetTextbox("Posting Date","GOHEAD-BUDAT","",ConvertDate(DT_MIGO_0110_POSTING_DATE_OCC2),False)
Call ClickButtonIfExist("MIGO_OK_GO",False)

CAll SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,DT_TRUE,False)
Call SetTableData("SAPLMIGOTV_GOITEM","OK","1","","",DT_TRUE,False)
Call PressEnter()  

Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Call VerifyifGuiLabelExists(DT_MIGO_120_CHECK_TEXT_DELIVERY_NOTE_ALREADY_RECEIVED)
Call ClickButton("Continue   \(Enter\)",False)

Call LogOff()
Call FinalStatus ()
